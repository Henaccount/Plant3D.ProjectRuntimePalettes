namespace Plant3D.ProjectRuntimePalettes.Models;

public sealed class ProjectPaletteModel
{
    public ProjectPaletteModel(
        ProjectRuntimeContext context,
        IReadOnlyList<ProjectPaletteTreeNode> rootNodes,
        IReadOnlyList<ProjectPaletteItem> allItems)
    {
        Context = context;
        RootNodes = rootNodes;
        AllItems = allItems;
    }

    public ProjectRuntimeContext Context { get; }

    public IReadOnlyList<ProjectPaletteTreeNode> RootNodes { get; }

    public IReadOnlyList<ProjectPaletteItem> AllItems { get; }

    public int TotalItems => AllItems.Count;

    public string ModeText
    {
        get
        {
            if (Context.EffectiveRespectSupportedStandards)
            {
                return $"SupportedStandards filter: ON ({Context.CurrentStandardToken})";
            }

            if (Context.RequestedRespectSupportedStandards)
            {
                return "SupportedStandards filter requested, but project standard could not be resolved";
            }

            return "SupportedStandards filter: OFF";
        }
    }
}
