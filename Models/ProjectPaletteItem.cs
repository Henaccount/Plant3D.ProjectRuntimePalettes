namespace Plant3D.ProjectRuntimePalettes.Models;

public sealed record ProjectPaletteItem(
    PaletteCategory Category,
    string ClassName,
    string DisplayName,
    string VisualName,
    string? SymbolName,
    string? LineStyleName,
    int? SupportedStandardsMask,
    string SourcePath,
    string? ParentClassName,
    IReadOnlyList<string> StyleCandidates,
    bool? TpIncluded,
    bool IsLeafClass = true)
{
    public string UniqueKey => $"{Category}|{ClassName}|{VisualName}|{SourcePath}";
}
