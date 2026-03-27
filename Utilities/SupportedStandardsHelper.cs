using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Data.Sqlite;

namespace Plant3D.ProjectRuntimePalettes.Utilities;

public static partial class SupportedStandardsHelper
{
    public const int PipMask = 3;
    public const int IsaMask = 5;
    public const int IsoMask = 9;
    public const int DinMask = 17;
    public const int JisIsoMask = 33;

    private static readonly IReadOnlyDictionary<string, int> TokenToMask = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        ["PIP"] = PipMask,
        ["ISA"] = IsaMask,
        ["ISO"] = IsoMask,
        ["DIN"] = DinMask,
        ["JIS-ISO"] = JisIsoMask,
        ["JIS ISO"] = JisIsoMask,
        ["JISISO"] = JisIsoMask
    };

    private static readonly string[] TokensInPriorityOrder =
    [
        "JIS-ISO",
        "PIP",
        "ISA",
        "DIN",
        "ISO"
    ];

    [GeneratedRegex("[^A-Z0-9]+")]
    private static partial Regex NonAlphaNumericRegex();

    public static bool MatchesProjectStandard(int supportedMask, int currentProjectMask)
    {
        if (currentProjectMask == 0)
        {
            return false;
        }

        if ((supportedMask & currentProjectMask) == currentProjectMask)
        {
            return true;
        }

        var projectBit = GetStandardBit(currentProjectMask);
        return projectBit != 0 && (supportedMask & projectBit) == projectBit;
    }

    public static (string? Token, int Mask) ResolveCurrentProjectStandard(string? pnIdPartXmlPath, string processPowerDcfPath)
    {
        if (!string.IsNullOrWhiteSpace(pnIdPartXmlPath) && File.Exists(pnIdPartXmlPath)
            && TryResolveCurrentProjectStandardFromXml(pnIdPartXmlPath, out var xmlToken))
        {
            return (xmlToken, TokenToMask[xmlToken!]);
        }

        if (File.Exists(processPowerDcfPath) && TryResolveCurrentProjectStandardFromDcf(processPowerDcfPath, out var dcfToken))
        {
            return (dcfToken, TokenToMask[dcfToken!]);
        }

        return (null, 0);
    }

    public static int? ParseSupportedStandards(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        if (int.TryParse(raw.Trim(), out var integerMask))
        {
            return integerMask;
        }

        var normalized = NormalizeStandardText(raw);
        var combined = 0;

        foreach (var token in TokensInPriorityOrder)
        {
            var normalizedToken = NormalizeStandardText(token);
            if (normalized.Contains(normalizedToken, StringComparison.Ordinal))
            {
                combined |= TokenToMask[token];
            }
        }

        return combined == 0 ? null : combined;
    }

    public static bool TryResolveCurrentProjectStandardFromXml(string xmlPath, out string? token)
    {
        token = null;

        try
        {
            var document = XDocument.Load(xmlPath, LoadOptions.None);

            var projectStandardType = document
                .Descendants()
                .FirstOrDefault(element => string.Equals(element.Name.LocalName, "ProjectStandardType", StringComparison.OrdinalIgnoreCase))
                ?.Value;

            if (TryParseStandardToken(projectStandardType, out token))
            {
                return true;
            }

            foreach (var attr in document.Descendants().Attributes())
            {
                if (LooksLikeStandardField(attr.Name.LocalName) && TryParseStandardToken(attr.Value, out token))
                {
                    return true;
                }
            }

            foreach (var element in document.Descendants())
            {
                if (LooksLikeStandardField(element.Name.LocalName) && TryParseStandardToken(element.Value, out token))
                {
                    return true;
                }
            }
        }
        catch
        {
        }

        return false;
    }

    public static bool TryResolveCurrentProjectStandardFromDcf(string dcfPath, out string? token)
    {
        token = null;

        try
        {
            using var connection = new SqliteConnection(new SqliteConnectionStringBuilder
            {
                DataSource = dcfPath,
                Mode = SqliteOpenMode.ReadOnly
            }.ToString());
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = "SELECT Project_Standard FROM PnPProject LIMIT 1";
            var raw = command.ExecuteScalar() as string;

            return TryParseStandardToken(raw, out token);
        }
        catch
        {
            return false;
        }
    }

    public static bool TryParseStandardToken(string? raw, out string? token)
    {
        token = null;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        var normalized = NormalizeStandardText(raw);
        foreach (var candidate in TokensInPriorityOrder)
        {
            if (normalized.Contains(NormalizeStandardText(candidate), StringComparison.Ordinal))
            {
                token = candidate;
                return true;
            }
        }

        return false;
    }

    public static int GetStandardBit(int currentProjectMask)
    {
        return currentProjectMask switch
        {
            PipMask => 2,
            IsaMask => 4,
            IsoMask => 8,
            DinMask => 16,
            JisIsoMask => 32,
            _ => 0
        };
    }

    private static string NormalizeStandardText(string raw)
    {
        var upper = raw.Trim().ToUpperInvariant();
        var normalized = NonAlphaNumericRegex().Replace(upper, " ").Trim();
        return normalized;
    }

    private static bool LooksLikeStandardField(string localName)
    {
        var normalized = SearchText.Normalize(localName);
        return normalized.Contains("standard", StringComparison.Ordinal)
            || normalized.Contains("symbology", StringComparison.Ordinal)
            || normalized.Contains("tool palette", StringComparison.Ordinal)
            || normalized.Contains("palette", StringComparison.Ordinal);
    }
}
