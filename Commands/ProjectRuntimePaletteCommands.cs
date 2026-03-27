using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Plant3D.ProjectRuntimePalettes.Services;
using Plant3D.ProjectRuntimePalettes.UI;

namespace Plant3D.ProjectRuntimePalettes.Commands;

public sealed class ProjectRuntimePaletteCommands
{
    private static readonly ProjectContextLocator ContextLocator = new();
    private static readonly DcfProjectModelReader ModelReader = new();
    private static readonly ProjectStyleLibraryService StyleLibraryService = new();
    private static readonly SymbolPreviewService PreviewService = new(StyleLibraryService);
    private static readonly ToolExecutionService ExecutionService = new(StyleLibraryService);
    private static readonly ProjectRuntimePaletteHost PaletteHost = new(PreviewService, ExecutionService);

    [CommandMethod("PROJECTPIDPALETTESINSERT")]
    public void ExecutePendingProjectPidPaletteInsertion()
    {
        ExecutionService.ExecutePending();
    }

    [CommandMethod("PROJECTPIDPALETTES")]
    public void ShowProjectPidPalettes()
    {
        var document = AcadApp.DocumentManager.MdiActiveDocument;
        if (document is null)
        {
            return;
        }

        var editor = document.Editor;
        var respectSupportedStandards = PromptForMode(editor);
        if (respectSupportedStandards is null)
        {
            return;
        }

        if (!ContextLocator.TryLocate(respectSupportedStandards.Value, editor, out var context, out var message))
        {
            if (!string.IsNullOrWhiteSpace(message))
            {
                editor.WriteMessage($"\n{message}");
            }

            return;
        }

        if (!string.IsNullOrWhiteSpace(message))
        {
            editor.WriteMessage($"\n{message}");
        }

        try
        {
            var model = ModelReader.Read(context, editor);
            PaletteHost.Show(model);
            editor.WriteMessage($"\nProject P&ID palette loaded with {model.TotalItems} styled leaf class(es). Insertion now uses the Plant P&ID runtime API directly, without pre-existing palette-definition dependencies.");
        }
        catch (System.Exception ex)
        {
            editor.WriteMessage($"\nFailed to build the project P&ID palette: {ex.Message}");
        }
    }

    private static bool? PromptForMode(Editor editor)
    {
        var prompt = new PromptKeywordOptions("\nShow project P&ID palette mode [All/RespectSupportedStandards] <All>: ")
        {
            AllowNone = true
        };
        prompt.Keywords.Add("All");
        prompt.Keywords.Add("RespectSupportedStandards");

        var result = editor.GetKeywords(prompt);
        if (result.Status == PromptStatus.Cancel)
        {
            return null;
        }

        return string.Equals(result.StringResult, "RespectSupportedStandards", StringComparison.OrdinalIgnoreCase);
    }
}
