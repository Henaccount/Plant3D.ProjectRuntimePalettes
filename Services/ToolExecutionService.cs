using System.Threading.Tasks;
using Autodesk.AutoCAD.EditorInput;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using AcadGeom = Autodesk.AutoCAD.Geometry;
using Plant3D.ProjectRuntimePalettes.Models;

namespace Plant3D.ProjectRuntimePalettes.Services;

public sealed class ToolExecutionService
{
    private const string PendingInsertionCommandName = "PROJECTPIDPALETTESINSERT";

    private readonly ProjectStyleLibraryService _styleLibraryService;
    private readonly PidRuntimeApi _pidRuntimeApi = new();
    private readonly object _pendingSyncRoot = new();
    private PendingInsertionRequest? _pendingRequest;

    public ToolExecutionService(ProjectStyleLibraryService styleLibraryService)
    {
        _styleLibraryService = styleLibraryService;
    }

    public bool CanExecute(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var styleResolution = _styleLibraryService.Resolve(item, context);
        var styleName = ResolveInsertionStyleName(item, styleResolution);
        return !string.IsNullOrWhiteSpace(styleName);
    }

    public string GetFailureReason(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var styleResolution = _styleLibraryService.Resolve(item, context);
        return BuildFailureMessage(item, styleResolution);
    }

    public bool QueueExecute(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var document = AcadApp.DocumentManager.MdiActiveDocument;
        if (document is null)
        {
            return false;
        }

        lock (_pendingSyncRoot)
        {
            _pendingRequest = new PendingInsertionRequest(item, context);
        }

        try
        {
            AcadApp.DocumentManager.ExecuteInCommandContextAsync(
                async _ =>
                {
                    ExecutePending();
                    await Task.CompletedTask;
                },
                null);
            return true;
        }
        catch
        {
            document.SendStringToExecute(PendingInsertionCommandName + " ", true, false, false);
            return true;
        }
    }

    public bool Execute(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        return QueueExecute(item, context);
    }

    public void ExecutePending()
    {
        PendingInsertionRequest? request;
        lock (_pendingSyncRoot)
        {
            request = _pendingRequest;
            _pendingRequest = null;
        }

        if (request is null)
        {
            WriteMessage("No pending palette item insertion request is available.");
            return;
        }

        ExecuteNow(request.Item, request.Context);
    }

    private bool ExecuteNow(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var document = AcadApp.DocumentManager.MdiActiveDocument;
        if (document is null)
        {
            return false;
        }

        var editor = document.Editor;
        var database = document.Database;
        var styleResolution = _styleLibraryService.Resolve(item, context);
        var styleName = ResolveInsertionStyleName(item, styleResolution);

        if (string.IsNullOrWhiteSpace(styleName))
        {
            WriteMessage(BuildFailureMessage(item, styleResolution));
            return false;
        }

        if (!_pidRuntimeApi.TryEnsureStyleLoaded(styleName, context.SymbolStyleDrawingPath, database, out var styleId, out var styleMessage))
        {
            WriteMessage($"Cannot insert '{item.DisplayName}'. {styleMessage}");
            return false;
        }

        return item.Category == PaletteCategory.Lines
            ? InsertLine(item, styleResolution, styleName, styleId, editor, database)
            : InsertAsset(item, styleResolution, styleName, styleId, editor, database);
    }

    private bool InsertAsset(
        ProjectPaletteItem item,
        ProjectStyleResolution styleResolution,
        string styleName,
        AcadDb.ObjectId styleId,
        Editor editor,
        AcadDb.Database database)
    {
        var pointResult = editor.GetPoint($"\nSpecify insertion point for {item.DisplayName}: ");
        if (pointResult.Status != PromptStatus.OK)
        {
            return false;
        }

        try
        {
            if (_pidRuntimeApi.TryInsertAsset(item.ClassName, styleName, styleId, pointResult.Value, database, out var insertMessage))
            {
                return true;
            }

            WriteMessage($"Cannot insert '{item.DisplayName}'. {BuildInsertionFailureSuffix(styleResolution, styleName, insertMessage)}");
            return false;
        }
        catch (System.Exception ex)
        {
            WriteMessage($"Cannot insert '{item.DisplayName}'. {BuildInsertionFailureSuffix(styleResolution, styleName, ex.Message)}");
            return false;
        }
    }

    private bool InsertLine(
        ProjectPaletteItem item,
        ProjectStyleResolution styleResolution,
        string styleName,
        AcadDb.ObjectId styleId,
        Editor editor,
        AcadDb.Database database)
    {
        var vertices = PromptForLineVertices(editor, item.DisplayName);
        if (vertices is null || vertices.Count < 2)
        {
            return false;
        }

        try
        {
            if (_pidRuntimeApi.TryInsertLine(item.ClassName, styleName, styleId, vertices, database, out var insertMessage))
            {
                return true;
            }

            WriteMessage($"Cannot insert '{item.DisplayName}'. {BuildInsertionFailureSuffix(styleResolution, styleName, insertMessage)}");
            return false;
        }
        catch (System.Exception ex)
        {
            WriteMessage($"Cannot insert '{item.DisplayName}'. {BuildInsertionFailureSuffix(styleResolution, styleName, ex.Message)}");
            return false;
        }
    }

    private static AcadGeom.Point3dCollection? PromptForLineVertices(Editor editor, string displayName)
    {
        var vertices = new AcadGeom.Point3dCollection();

        var firstPoint = editor.GetPoint($"\nSpecify start point for {displayName}: ");
        if (firstPoint.Status != PromptStatus.OK)
        {
            return null;
        }

        vertices.Add(firstPoint.Value);

        while (true)
        {
            var options = new PromptPointOptions("\nSpecify next point or press Enter to finish: ")
            {
                AllowNone = true,
                UseBasePoint = true,
                BasePoint = vertices[vertices.Count - 1]
            };

            var nextPoint = editor.GetPoint(options);
            if (nextPoint.Status == PromptStatus.None)
            {
                return vertices.Count >= 2 ? vertices : null;
            }

            if (nextPoint.Status != PromptStatus.OK)
            {
                return null;
            }

            vertices.Add(nextPoint.Value);
        }
    }

    private string BuildFailureMessage(ProjectPaletteItem item, ProjectStyleResolution styleResolution)
    {
        var styleName = ResolveInsertionStyleName(item, styleResolution);
        if (string.IsNullOrWhiteSpace(styleName))
        {
            return item.StyleCandidates.Count > 0
                ? $"Cannot insert '{item.DisplayName}'. No project style could be resolved from: {string.Join(", ", item.StyleCandidates)}."
                : $"Cannot insert '{item.DisplayName}'. No project style is connected to this class.";
        }

        if (!_pidRuntimeApi.IsAvailable)
        {
            return $"Cannot insert '{item.DisplayName}'. Style '{styleName}' was resolved, but the Plant P&ID runtime API is unavailable. {_pidRuntimeApi.AvailabilityMessage}";
        }

        return item.Category == PaletteCategory.Lines
            ? $"Cannot insert '{item.DisplayName}'. Line style '{styleName}' is available, but line creation still failed in the Plant P&ID runtime API."
            : $"Cannot insert '{item.DisplayName}'. Style '{styleName}'{BuildBlockClause(styleResolution)} is available, but asset creation still failed in the Plant P&ID runtime API.";
    }

    private static string ResolveInsertionStyleName(ProjectPaletteItem item, ProjectStyleResolution styleResolution)
    {
        if (!string.IsNullOrWhiteSpace(styleResolution.StyleInfo?.StyleName))
        {
            return styleResolution.StyleInfo.StyleName;
        }

        return item.StyleCandidates.FirstOrDefault(candidate => !string.IsNullOrWhiteSpace(candidate))
            ?? item.VisualName
            ?? string.Empty;
    }

    private static string BuildInsertionFailureSuffix(ProjectStyleResolution styleResolution, string styleName, string insertMessage)
    {
        var blockClause = BuildBlockClause(styleResolution);
        return $"Style '{styleName}'{blockClause} could not be inserted directly through the Plant P&ID runtime API. {insertMessage}";
    }

    private static string BuildBlockClause(ProjectStyleResolution styleResolution)
    {
        return !string.IsNullOrWhiteSpace(styleResolution.DisplayBlockName)
            ? $" (block '{styleResolution.DisplayBlockName}')"
            : string.Empty;
    }

    private static void WriteMessage(string message)
    {
        var editor = AcadApp.DocumentManager.MdiActiveDocument?.Editor;
        editor?.WriteMessage($"{message}");
    }

    private sealed record PendingInsertionRequest(ProjectPaletteItem Item, ProjectRuntimeContext Context);
}
