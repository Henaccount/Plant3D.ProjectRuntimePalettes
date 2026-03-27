namespace Plant3D.ProjectRuntimePalettes.Models;

public sealed record ProjectRuntimeContext(
    string ProjectName,
    string ProjectRoot,
    string ProcessPowerDcfPath,
    string? PnIdPartXmlPath,
    string? SymbolStyleDrawingPath,
    bool RequestedRespectSupportedStandards,
    bool EffectiveRespectSupportedStandards,
    string? CurrentStandardToken,
    int CurrentStandardMask);
