using System.Drawing;
using Autodesk.AutoCAD.Windows;
using Plant3D.ProjectRuntimePalettes.Models;
using Plant3D.ProjectRuntimePalettes.Services;

namespace Plant3D.ProjectRuntimePalettes.UI;

public sealed class ProjectRuntimePaletteHost
{
    private static readonly Guid PaletteGuid = new("3D2B2ED8-1F56-4D8B-97A0-72B0A9CF6B6D");

    private readonly SymbolPreviewService _previewService;
    private readonly ToolExecutionService _executionService;
    private PaletteSet? _paletteSet;
    private ProjectRuntimePaletteControl? _control;

    public ProjectRuntimePaletteHost(SymbolPreviewService previewService, ToolExecutionService executionService)
    {
        _previewService = previewService;
        _executionService = executionService;
    }

    public void Show(ProjectPaletteModel model)
    {
        EnsurePaletteCreated();
        _control!.LoadModel(model);
        _paletteSet!.Visible = true;
    }

    private void EnsurePaletteCreated()
    {
        if (_paletteSet is not null)
        {
            return;
        }

        _paletteSet = new PaletteSet("Project P&ID Tool Palette", PaletteGuid)
        {
            DockEnabled = DockSides.Left | DockSides.Right,
            Size = new Size(430, 700)
        };

        _control = new ProjectRuntimePaletteControl(_previewService, _executionService);
        _paletteSet.Add("Project", _control);
    }
}
