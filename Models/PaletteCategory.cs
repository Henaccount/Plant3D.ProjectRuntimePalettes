using System.Xml.Linq;
using Plant3D.ProjectRuntimePalettes.Utilities;

namespace Plant3D.ProjectRuntimePalettes.Models;

public enum PaletteCategory
{
    Equipment,
    Valves,
    Fittings,
    Speciality,
    Reducers,
    Instrumentation,
    Lines,
    Nozzles,
    NonEngineeringItems
}

public static class PaletteCategoryInfo
{
    private static readonly IReadOnlyDictionary<PaletteCategory, string> DisplayNames = new Dictionary<PaletteCategory, string>
    {
        [PaletteCategory.Equipment] = "Equipment",
        [PaletteCategory.Valves] = "Valves",
        [PaletteCategory.Fittings] = "Fittings",
        [PaletteCategory.Speciality] = "Speciality",
        [PaletteCategory.Reducers] = "Reducers",
        [PaletteCategory.Instrumentation] = "Instrumentation",
        [PaletteCategory.Lines] = "Lines",
        [PaletteCategory.Nozzles] = "Nozzles",
        [PaletteCategory.NonEngineeringItems] = "NonEngineeringItems"
    };

    private static readonly (PaletteCategory Category, string[] Terms)[] MatchOrder =
    [
        (PaletteCategory.NonEngineeringItems, ["non engineering", "nonengineering", "flow arrow", "line breaker", "connector", "annotation", "gap"]),
        (PaletteCategory.Valves, ["valves", "valve"]),
        (PaletteCategory.Reducers, ["reducers", "reducer"]),
        (PaletteCategory.Fittings, ["fittings", "fitting"]),
        (PaletteCategory.Speciality, ["speciality", "specialty"]),
        (PaletteCategory.Nozzles, ["nozzles", "nozzle"]),
        (PaletteCategory.Instrumentation, ["instrumentation", "instruments", "instrument"]),
        (PaletteCategory.Equipment, ["equipment"]),
        (PaletteCategory.Lines, ["signal lines", "signal line", "pipe line groups", "pipe line group", "pipe lines", "pipe line", "lines", "line"])
    ];

    public static IReadOnlyList<PaletteCategory> OrderedCategories { get; } =
    [
        PaletteCategory.Equipment,
        PaletteCategory.Valves,
        PaletteCategory.Fittings,
        PaletteCategory.Speciality,
        PaletteCategory.Reducers,
        PaletteCategory.Instrumentation,
        PaletteCategory.Lines,
        PaletteCategory.Nozzles,
        PaletteCategory.NonEngineeringItems
    ];

    public static string GetDisplayName(PaletteCategory category) => DisplayNames[category];

    public static int GetSortKey(PaletteCategory category)
    {
        for (var i = 0; i < OrderedCategories.Count; i++)
        {
            if (OrderedCategories[i] == category)
            {
                return i;
            }
        }

        return int.MaxValue;
    }


    public static bool TryResolveCategory(XElement element, string? className, string? displayName, out PaletteCategory category)
    {
        var bag = new List<string>();

        foreach (var node in element.AncestorsAndSelf().Take(6))
        {
            bag.Add(node.Name.LocalName);
            bag.AddRange(XmlReadHelpers.ReadDirectIdentityValues(node));
        }

        if (!string.IsNullOrWhiteSpace(className))
        {
            bag.Add(className);
        }

        if (!string.IsNullOrWhiteSpace(displayName))
        {
            bag.Add(displayName);
        }

        var normalized = SearchText.Normalize(string.Join(' ', bag));

        foreach (var entry in MatchOrder)
        {
            if (entry.Terms.Any(term => SearchText.ContainsPhrase(normalized, term)))
            {
                category = entry.Category;
                return true;
            }
        }

        category = default;
        return false;
    }
}
