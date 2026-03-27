namespace Plant3D.ProjectRuntimePalettes.Models;

public sealed class ProjectPaletteTreeNode
{
    public ProjectPaletteTreeNode(
        string nodeKey,
        PaletteCategory category,
        string displayName,
        string? className,
        bool isSynthetic)
    {
        NodeKey = nodeKey;
        Category = category;
        DisplayName = displayName;
        ClassName = className;
        IsSynthetic = isSynthetic;
    }

    public string NodeKey { get; }

    public PaletteCategory Category { get; }

    public string DisplayName { get; }

    public string? ClassName { get; }

    public bool IsSynthetic { get; }

    public ProjectPaletteItem? PaletteItem { get; set; }

    public List<ProjectPaletteTreeNode> Children { get; } = new();

    public IReadOnlyList<ProjectPaletteItem> DescendantItems { get; internal set; } = Array.Empty<ProjectPaletteItem>();

    public bool IsInsertable => PaletteItem is not null;
}
