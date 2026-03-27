using Autodesk.AutoCAD.EditorInput;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Plant3D.ProjectRuntimePalettes.Models;
using Plant3D.ProjectRuntimePalettes.Utilities;

namespace Plant3D.ProjectRuntimePalettes.Services;

public sealed class ProjectContextLocator
{
    private static readonly string[] ProjectPathPropertyCandidates =
    [
        "ProjectFolderPath",
        "ProjectPath",
        "ProjectDirectory",
        "ProjectFolder",
        "RootFolder",
        "HomeFolder",
        "ProjectHomePath"
    ];

    private static readonly string[] PartPathPropertyCandidates =
    [
        "ProjectFolderPath",
        "ProjectPath",
        "ProjectDirectory",
        "Location",
        "Path",
        "AbsoluteFileName",
        "DataFileName"
    ];

    public bool TryLocate(
        bool requestedRespectSupportedStandards,
        Editor editor,
        out ProjectRuntimeContext context,
        out string message)
    {
        context = null!;
        message = string.Empty;

        var currentProject = PlantApplication.CurrentProject;
        if (currentProject is null)
        {
            message = "No active Plant 3D project is available.";
            return false;
        }

        var projectRoot = ResolveProjectRoot(currentProject);
        if (string.IsNullOrWhiteSpace(projectRoot) || !Directory.Exists(projectRoot))
        {
            message = "Unable to resolve the Plant project root folder.";
            return false;
        }

        var processPowerDcfPath = LocateProjectFile(projectRoot, "ProcessPower.dcf");
        if (processPowerDcfPath is null)
        {
            message = "Could not locate ProcessPower.dcf below the Plant project root.";
            return false;
        }

        var pnIdPartXmlPath = LocatePnIdPartXml(projectRoot);
        var symbolStyleDrawingPath = LocateProjectFile(projectRoot, "projSymbolStyle.dwg");

        var (standardToken, standardMask) = SupportedStandardsHelper.ResolveCurrentProjectStandard(pnIdPartXmlPath, processPowerDcfPath);
        var effectiveRespectSupportedStandards = requestedRespectSupportedStandards && standardMask != 0;

        if (requestedRespectSupportedStandards && standardMask == 0)
        {
            message = "Could not resolve the current P&ID standard from the PnIdPart XML or ProcessPower.dcf. The palette will be shown without SupportedStandards filtering.";
        }

        var projectName = TryGetStringProperty(currentProject, "ProjectName")
            ?? Path.GetFileName(projectRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));

        context = new ProjectRuntimeContext(
            projectName,
            projectRoot,
            processPowerDcfPath,
            pnIdPartXmlPath,
            symbolStyleDrawingPath,
            requestedRespectSupportedStandards,
            effectiveRespectSupportedStandards,
            standardToken,
            standardMask);

        return true;
    }

    private static string? ResolveProjectRoot(PlantProject currentProject)
    {
        foreach (var propertyName in ProjectPathPropertyCandidates)
        {
            var value = TryGetStringProperty(currentProject, propertyName);
            var normalized = NormalizePathCandidate(value);
            if (normalized is not null)
            {
                return normalized;
            }
        }

        foreach (Project part in currentProject.ProjectParts)
        {
            foreach (var propertyName in PartPathPropertyCandidates)
            {
                var value = TryGetStringProperty(part, propertyName);
                var normalized = NormalizePathCandidate(value);
                if (normalized is not null)
                {
                    return normalized;
                }
            }
        }

        var activeDocumentPath = AcadApp.DocumentManager.MdiActiveDocument?.Name;
        if (!string.IsNullOrWhiteSpace(activeDocumentPath))
        {
            var currentDirectory = Directory.Exists(activeDocumentPath)
                ? new DirectoryInfo(activeDocumentPath)
                : new FileInfo(activeDocumentPath).Directory;

            while (currentDirectory is not null)
            {
                if (File.Exists(Path.Combine(currentDirectory.FullName, "ProcessPower.dcf")))
                {
                    return currentDirectory.FullName;
                }

                currentDirectory = currentDirectory.Parent;
            }
        }

        return null;
    }

    private static string? LocateProjectFile(string projectRoot, string fileName)
    {
        var directPath = Path.Combine(projectRoot, fileName);
        if (File.Exists(directPath))
        {
            return directPath;
        }

        try
        {
            return Directory.EnumerateFiles(projectRoot, fileName, SearchOption.AllDirectories).FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }

    private static string? LocatePnIdPartXml(string projectRoot)
    {
        var directCandidates = Directory
            .EnumerateFiles(projectRoot, "*PnIdPart.xml", SearchOption.TopDirectoryOnly)
            .Concat(Directory.EnumerateFiles(projectRoot, "*.PnIdPart.xml", SearchOption.TopDirectoryOnly))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (directCandidates.Count > 0)
        {
            return directCandidates[0];
        }

        try
        {
            return Directory
                .EnumerateFiles(projectRoot, "*PnIdPart.xml", SearchOption.AllDirectories)
                .Concat(Directory.EnumerateFiles(projectRoot, "*.PnIdPart.xml", SearchOption.AllDirectories))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }
        catch
        {
            return null;
        }
    }

    private static string? NormalizePathCandidate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        try
        {
            if (Directory.Exists(value))
            {
                return Path.GetFullPath(value);
            }

            if (File.Exists(value))
            {
                return Path.GetDirectoryName(Path.GetFullPath(value));
            }
        }
        catch
        {
        }

        return null;
    }

    private static string? TryGetStringProperty(object instance, string propertyName)
    {
        var property = instance.GetType().GetProperty(propertyName);
        if (property is null)
        {
            return null;
        }

        var value = property.GetValue(instance);
        return value switch
        {
            null => null,
            string text => text,
            _ => value.ToString()
        };
    }
}
