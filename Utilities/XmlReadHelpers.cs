using System.Xml.Linq;

namespace Plant3D.ProjectRuntimePalettes.Utilities;

public static class XmlReadHelpers
{
    private static readonly HashSet<string> MetaAttributeNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "name",
        "key",
        "id",
        "property",
        "type",
        "category"
    };

    public static IEnumerable<string> ReadDirectIdentityValues(XElement element)
    {
        foreach (var attr in element.Attributes())
        {
            if (attr.IsNamespaceDeclaration)
            {
                continue;
            }

            if (IsIdentityName(attr.Name.LocalName))
            {
                var value = attr.Value.Trim();
                if (!string.IsNullOrWhiteSpace(value))
                {
                    yield return value;
                }
            }
        }

        if (!element.HasElements)
        {
            var value = element.Value.Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                yield return value;
            }
        }
    }

    public static string? ReadNearValue(XElement root, int maxDepth, params string[] keys)
    {
        return ReadNearValues(root, maxDepth, keys).FirstOrDefault();
    }

    public static IReadOnlyList<string> ReadNearValues(XElement root, int maxDepth, params string[] keys)
    {
        var normalizedKeys = keys.Select(SearchText.Normalize).ToHashSet(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var values = new List<string>();

        foreach (var element in EnumerateWithinDepth(root, maxDepth))
        {
            foreach (var attr in element.Attributes())
            {
                if (attr.IsNamespaceDeclaration)
                {
                    continue;
                }

                if (normalizedKeys.Contains(SearchText.Normalize(attr.Name.LocalName)))
                {
                    AddValue(attr.Value);
                }
            }

            if (normalizedKeys.Contains(SearchText.Normalize(element.Name.LocalName)))
            {
                foreach (var value in ExtractElementValues(element))
                {
                    AddValue(value);
                }
            }

            foreach (var markerAttribute in element.Attributes())
            {
                if (markerAttribute.IsNamespaceDeclaration)
                {
                    continue;
                }

                if (MetaAttributeNames.Contains(markerAttribute.Name.LocalName)
                    && normalizedKeys.Contains(SearchText.Normalize(markerAttribute.Value)))
                {
                    foreach (var value in ExtractElementValues(element))
                    {
                        AddValue(value);
                    }
                }
            }
        }

        return values;

        void AddValue(string? value)
        {
            var trimmed = value?.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return;
            }

            if (seen.Add(trimmed))
            {
                values.Add(trimmed);
            }
        }
    }

    public static string GetPseudoPath(XElement element)
    {
        return "/" + string.Join("/", element.AncestorsAndSelf().Reverse().Select(node => node.Name.LocalName));
    }

    private static IEnumerable<XElement> EnumerateWithinDepth(XElement root, int maxDepth)
    {
        yield return root;

        foreach (var descendant in root.Descendants())
        {
            var depth = descendant.Ancestors().TakeWhile(ancestor => ancestor != root).Count() + 1;
            if (depth <= maxDepth)
            {
                yield return descendant;
            }
        }
    }

    private static IEnumerable<string> ExtractElementValues(XElement element)
    {
        foreach (var attr in element.Attributes())
        {
            if (attr.IsNamespaceDeclaration || MetaAttributeNames.Contains(attr.Name.LocalName))
            {
                continue;
            }

            var value = attr.Value.Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                yield return value;
            }
        }

        var directText = element.Value.Trim();
        if (!string.IsNullOrWhiteSpace(directText) && !element.HasElements)
        {
            yield return directText;
        }

        foreach (var child in element.Elements())
        {
            if (SearchText.Normalize(child.Name.LocalName) is "value" or "text" or "displayname" or "symbolname" or "linestyle" or "linestylename")
            {
                var childValue = child.Value.Trim();
                if (!string.IsNullOrWhiteSpace(childValue))
                {
                    yield return childValue;
                }
            }
        }
    }

    private static bool IsIdentityName(string localName)
    {
        var normalized = SearchText.Normalize(localName);
        return normalized is "name"
            or "displayname"
            or "classname"
            or "description"
            or "title"
            or "category"
            or "group";
    }
}
