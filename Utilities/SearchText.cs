using System.Text;
using System.Text.RegularExpressions;

namespace Plant3D.ProjectRuntimePalettes.Utilities;

public static partial class SearchText
{
    [GeneratedRegex("\\s+")]
    private static partial Regex MultiSpaceRegex();

    public static string Normalize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var span = text.AsSpan();
        var sb = new StringBuilder(span.Length * 2);

        for (var i = 0; i < span.Length; i++)
        {
            var ch = span[i];
            if (!char.IsLetterOrDigit(ch))
            {
                AppendSpace(sb);
                continue;
            }

            if (i > 0 && NeedsWordBoundary(span[i - 1], ch, i + 1 < span.Length ? span[i + 1] : '\0'))
            {
                AppendSpace(sb);
            }

            sb.Append(char.ToLowerInvariant(ch));
        }

        return MultiSpaceRegex().Replace(sb.ToString(), " ").Trim();
    }

    public static string Compact(string? text)
    {
        var normalized = Normalize(text);
        return string.IsNullOrWhiteSpace(normalized)
            ? string.Empty
            : normalized.Replace(" ", string.Empty, StringComparison.Ordinal);
    }

    public static bool Matches(string normalizedHaystack, string? phrase)
    {
        return ContainsPhrase(normalizedHaystack, phrase) || ContainsAllTokens(normalizedHaystack, phrase);
    }

    public static bool ContainsPhrase(string normalizedHaystack, string? phrase)
    {
        var normalizedPhrase = Normalize(phrase);
        return !string.IsNullOrWhiteSpace(normalizedPhrase) && normalizedHaystack.Contains(normalizedPhrase, StringComparison.Ordinal);
    }

    public static bool ContainsAllTokens(string normalizedHaystack, string? phrase)
    {
        var normalizedPhrase = Normalize(phrase);
        if (string.IsNullOrWhiteSpace(normalizedPhrase))
        {
            return false;
        }

        return normalizedPhrase
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(token => token.Length > 1)
            .All(token => normalizedHaystack.Contains(token, StringComparison.Ordinal));
    }

    private static bool NeedsWordBoundary(char previous, char current, char next)
    {
        if (!char.IsLetterOrDigit(previous) || !char.IsLetterOrDigit(current))
        {
            return false;
        }

        if (char.IsDigit(previous) && char.IsLetter(current))
        {
            return true;
        }

        if (char.IsLetter(previous) && char.IsDigit(current))
        {
            return true;
        }

        if (char.IsUpper(current) && (char.IsLower(previous) || (char.IsUpper(previous) && char.IsLower(next))))
        {
            return true;
        }

        return false;
    }

    private static void AppendSpace(StringBuilder sb)
    {
        if (sb.Length > 0 && sb[^1] != ' ')
        {
            sb.Append(' ');
        }
    }
}
