using System.Xml.Linq;
using Autodesk.AutoCAD.EditorInput;
using Microsoft.Data.Sqlite;
using Plant3D.ProjectRuntimePalettes.Models;
using Plant3D.ProjectRuntimePalettes.Utilities;

namespace Plant3D.ProjectRuntimePalettes.Services;

public sealed class DcfProjectModelReader
{
    private static readonly IReadOnlyDictionary<PaletteCategory, string?> AliasRootClassNames = new Dictionary<PaletteCategory, string?>
    {
        [PaletteCategory.Equipment] = "Equipment",
        [PaletteCategory.Valves] = null,
        [PaletteCategory.Fittings] = "PipingFittings",
        [PaletteCategory.Speciality] = "PipingSpecialtyItems",
        [PaletteCategory.Reducers] = "Reducers",
        [PaletteCategory.Instrumentation] = "Instrumentation",
        [PaletteCategory.Lines] = "Lines",
        [PaletteCategory.Nozzles] = "Nozzles",
        [PaletteCategory.NonEngineeringItems] = "NonEngineeringItems"
    };

    public ProjectPaletteModel Read(ProjectRuntimeContext context, Editor? editor = null)
    {
        if (!File.Exists(context.ProcessPowerDcfPath))
        {
            throw new FileNotFoundException("ProcessPower.dcf was not found.", context.ProcessPowerDcfPath);
        }

        using var connection = OpenReadOnlyConnection(context.ProcessPowerDcfPath);
        var tables = ReadTables(connection);
        var attributes = ReadAttributes(connection);
        var properties = ReadProperties(connection);
        var columnAttributes = ReadColumnAttributes(connection);
        var xmlStyleCandidates = ReadClassStyleCandidatesFromPnIdPartXml(context.PnIdPartXmlPath);

        var candidates = BuildCandidates(tables, attributes, properties, columnAttributes, xmlStyleCandidates);
        var tpIncludedScope = candidates
            .Where(candidate => candidate.StyleCandidates.Count > 0)
            .ToList();
        var applyTpIncluded = ShouldApplyTpIncluded(tpIncludedScope);

        var actualNodes = new Dictionary<string, ProjectPaletteTreeNode>(StringComparer.OrdinalIgnoreCase);
        var categoryRoots = PaletteCategoryInfo.OrderedCategories.ToDictionary(category => category, _ => new List<ProjectPaletteTreeNode>());
        var allItems = new List<ProjectPaletteItem>();
        var nonLeafClassKeys = new HashSet<string>(
            candidates
                .Where(candidate => !string.IsNullOrWhiteSpace(candidate.BaseTable))
                .Select(candidate => BuildCategoryClassKey(candidate.BaseTable!, candidate.Category)),
            StringComparer.OrdinalIgnoreCase);

        foreach (var candidate in candidates.OrderBy(value => value.ClassName, StringComparer.OrdinalIgnoreCase))
        {
            var isVisibleInCurrentMode = !candidate.IsAbstract
                && candidate.StyleCandidates.Count > 0
                && IsVisibleInCurrentMode(candidate.SupportedStandardsMask, context)
                && (!applyTpIncluded || candidate.TpIncluded == true);

            var isLeafClass = !nonLeafClassKeys.Contains(BuildCategoryClassKey(candidate.ClassName, candidate.Category));

            ProjectPaletteItem? paletteItem = null;
            if (isVisibleInCurrentMode && isLeafClass)
            {
                var visualName = candidate.StyleCandidates[0];
                var symbolName = candidate.Category == PaletteCategory.Lines ? null : visualName;
                var lineStyleName = candidate.Category == PaletteCategory.Lines ? visualName : null;

                paletteItem = new ProjectPaletteItem(
                    candidate.Category,
                    candidate.ClassName,
                    candidate.DisplayName,
                    visualName,
                    symbolName,
                    lineStyleName,
                    candidate.SupportedStandardsMask,
                    $"ProcessPower.dcf/PnPTables/{candidate.ClassName}",
                    candidate.BaseTable,
                    candidate.StyleCandidates,
                    candidate.TpIncluded,
                    isLeafClass);

                allItems.Add(paletteItem);
            }

            actualNodes[candidate.ClassName] = new ProjectPaletteTreeNode(
                candidate.ClassName,
                candidate.Category,
                candidate.DisplayName,
                candidate.ClassName,
                isSynthetic: false)
            {
                PaletteItem = paletteItem
            };
        }

        foreach (var candidate in candidates.OrderBy(value => value.ClassName, StringComparer.OrdinalIgnoreCase))
        {
            var node = actualNodes[candidate.ClassName];
            if (!string.IsNullOrWhiteSpace(candidate.BaseTable)
                && actualNodes.TryGetValue(candidate.BaseTable, out var parentNode)
                && parentNode.Category == node.Category)
            {
                parentNode.Children.Add(node);
            }
            else
            {
                categoryRoots[candidate.Category].Add(node);
            }
        }

        foreach (var category in PaletteCategoryInfo.OrderedCategories)
        {
            foreach (var rootNode in categoryRoots[category])
            {
                SortChildren(rootNode);
                PopulateDescendantItems(rootNode, pruneEmptyChildren: true);
            }
        }

        var syntheticRoots = new List<ProjectPaletteTreeNode>();
        foreach (var category in PaletteCategoryInfo.OrderedCategories)
        {
            var displayName = PaletteCategoryInfo.GetDisplayName(category);
            var syntheticRoot = new ProjectPaletteTreeNode($"__root__{category}", category, displayName, AliasRootClassNames[category], isSynthetic: true);

            if (AliasRootClassNames[category] is { Length: > 0 } aliasClassName
                && actualNodes.TryGetValue(aliasClassName, out var aliasNode))
            {
                syntheticRoot.PaletteItem = aliasNode.PaletteItem;
                foreach (var child in aliasNode.Children)
                {
                    syntheticRoot.Children.Add(child);
                }

                foreach (var extraRoot in categoryRoots[category].Where(rootNode => !string.Equals(rootNode.ClassName, aliasClassName, StringComparison.OrdinalIgnoreCase)))
                {
                    syntheticRoot.Children.Add(extraRoot);
                }
            }
            else
            {
                foreach (var rootNode in categoryRoots[category])
                {
                    syntheticRoot.Children.Add(rootNode);
                }
            }

            SortChildren(syntheticRoot);
            syntheticRoot.DescendantItems = CollectDescendantItems(syntheticRoot);
            if (syntheticRoot.DescendantItems.Count > 0)
            {
                syntheticRoots.Add(syntheticRoot);
            }
        }

        var distinctItems = allItems
            .GroupBy(item => item.UniqueKey)
            .Select(group => group.First())
            .OrderBy(item => PaletteCategoryInfo.GetSortKey(item.Category))
            .ThenBy(item => item.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        editor?.WriteMessage(
            $"\nDiscovered {distinctItems.Count} style-bearing P&ID class(es) from ProcessPower.dcf"
            + (applyTpIncluded ? " with tpincluded filtering." : "."));

        return new ProjectPaletteModel(context, syntheticRoots, distinctItems);
    }

    private static List<TableCandidate> BuildCandidates(
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes,
        IReadOnlyDictionary<string, Dictionary<string, PropertyDefinition>> properties,
        IReadOnlyDictionary<string, Dictionary<string, Dictionary<string, string>>> columnAttributes,
        IReadOnlyDictionary<string, IReadOnlyList<string>> xmlStyleCandidates)
    {
        var candidates = new List<TableCandidate>();

        foreach (var table in tables.Values.OrderBy(value => value.TableName, StringComparer.OrdinalIgnoreCase))
        {
            if (!TryResolvePrimaryCategory(table.TableName, tables, out var category))
            {
                continue;
            }

            var displayName = ResolveDisplayName(table.TableName, tables, attributes);

            var styleCandidates = ResolveStyleCandidates(
                table.TableName,
                category,
                tables,
                attributes,
                properties,
                columnAttributes,
                xmlStyleCandidates);

            var supportedStandardsMask = SupportedStandardsHelper.ParseSupportedStandards(
                ResolveSupportedStandardsRaw(table.TableName, tables, attributes, properties, columnAttributes));

            bool? tpIncluded = null;
            if (TryResolveClassBooleanProperty(
                    table.TableName,
                    "tpincluded",
                    tables,
                    attributes,
                    properties,
                    columnAttributes,
                    out var resolvedTpIncluded))
            {
                tpIncluded = resolvedTpIncluded;
            }

            candidates.Add(new TableCandidate(
                table.TableName,
                table.BaseTable,
                displayName,
                category,
                table.IsAbstract,
                styleCandidates,
                supportedStandardsMask,
                tpIncluded));
        }

        return candidates;
    }

    private static bool ShouldApplyTpIncluded(IReadOnlyCollection<TableCandidate> candidates)
    {
        return candidates.Count > 0
            && candidates.All(candidate => candidate.TpIncluded.HasValue);
    }

    private static string BuildCategoryClassKey(string className, PaletteCategory category)
    {
        return $"{category}|{className}";
    }

    private static Dictionary<string, Dictionary<string, PropertyDefinition>> ReadProperties(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT TableName, PropertyName, PropertyType, DefaultValue
FROM PnPProperties";

        using var reader = command.ExecuteReader();
        var result = new Dictionary<string, Dictionary<string, PropertyDefinition>>(StringComparer.OrdinalIgnoreCase);

        while (reader.Read())
        {
            var tableName = reader.IsDBNull(0) ? null : reader.GetString(0);
            var propertyName = reader.IsDBNull(1) ? null : reader.GetString(1);
            var propertyType = reader.IsDBNull(2) ? null : reader.GetString(2);
            var defaultValue = reader.IsDBNull(3) ? null : reader.GetString(3);

            if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(propertyName))
            {
                continue;
            }

            if (!result.TryGetValue(tableName, out var tableProperties))
            {
                tableProperties = new Dictionary<string, PropertyDefinition>(StringComparer.OrdinalIgnoreCase);
                result[tableName] = tableProperties;
            }

            tableProperties[propertyName] = new PropertyDefinition(tableName, propertyName, propertyType, defaultValue);
        }

        return result;
    }

    private static Dictionary<string, Dictionary<string, Dictionary<string, string>>> ReadColumnAttributes(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT TableName, ColumnName, AttributeName, AttributeValue
FROM PnPColumnAttributes";

        using var reader = command.ExecuteReader();
        var result = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>(StringComparer.OrdinalIgnoreCase);

        while (reader.Read())
        {
            var tableName = reader.IsDBNull(0) ? null : reader.GetString(0);
            var columnName = reader.IsDBNull(1) ? null : reader.GetString(1);
            var attributeName = reader.IsDBNull(2) ? null : reader.GetString(2);
            var attributeValue = reader.IsDBNull(3) ? string.Empty : reader.GetString(3);

            if (string.IsNullOrWhiteSpace(tableName)
                || string.IsNullOrWhiteSpace(columnName)
                || string.IsNullOrWhiteSpace(attributeName))
            {
                continue;
            }

            if (!result.TryGetValue(tableName, out var tableColumns))
            {
                tableColumns = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                result[tableName] = tableColumns;
            }

            if (!tableColumns.TryGetValue(columnName, out var propertyAttributes))
            {
                propertyAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                tableColumns[columnName] = propertyAttributes;
            }

            propertyAttributes[attributeName] = attributeValue;
        }

        return result;
    }

    private static SqliteConnection OpenReadOnlyConnection(string dcfPath)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = dcfPath,
            Mode = SqliteOpenMode.ReadOnly
        };

        var connection = new SqliteConnection(builder.ToString());
        connection.Open();
        return connection;
    }

    private static Dictionary<string, TableRow> ReadTables(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT TableName, BaseTable, Abstract FROM PnPTables";

        using var reader = command.ExecuteReader();
        var result = new Dictionary<string, TableRow>(StringComparer.OrdinalIgnoreCase);

        while (reader.Read())
        {
            var tableName = reader.IsDBNull(0) ? null : reader.GetString(0);
            if (string.IsNullOrWhiteSpace(tableName))
            {
                continue;
            }

            var baseTable = reader.IsDBNull(1) ? null : reader.GetString(1);
            var abstractText = reader.IsDBNull(2) ? null : reader.GetString(2);
            result[tableName] = new TableRow(tableName, baseTable, ParseSqliteBoolean(abstractText));
        }

        return result;
    }

    private static Dictionary<string, Dictionary<string, string>> ReadAttributes(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT TableName, AttributeName, AttributeValue
FROM PnPTableAttributes";

        using var reader = command.ExecuteReader();
        var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        while (reader.Read())
        {
            var tableName = reader.IsDBNull(0) ? null : reader.GetString(0);
            var attributeName = reader.IsDBNull(1) ? null : reader.GetString(1);
            var attributeValue = reader.IsDBNull(2) ? string.Empty : reader.GetString(2);

            if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(attributeName))
            {
                continue;
            }

            if (!result.TryGetValue(tableName, out var tableAttributes))
            {
                tableAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                result[tableName] = tableAttributes;
            }

            tableAttributes[attributeName] = attributeValue;
        }

        return result;
    }

    private static Dictionary<string, IReadOnlyList<string>> ReadClassStyleCandidatesFromPnIdPartXml(string? pnIdPartXmlPath)
    {
        var result = new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(pnIdPartXmlPath) || !File.Exists(pnIdPartXmlPath))
        {
            return result;
        }

        try
        {
            var document = XDocument.Load(pnIdPartXmlPath, LoadOptions.None);

            foreach (var rule in document.Descendants().Where(element => string.Equals(element.Name.LocalName, "PnPRule", StringComparison.OrdinalIgnoreCase)))
            {
                var className = rule.Elements().FirstOrDefault(element => string.Equals(element.Name.LocalName, "Classname", StringComparison.OrdinalIgnoreCase))?.Value?.Trim();
                if (string.IsNullOrWhiteSpace(className))
                {
                    continue;
                }

                var orderedStyles = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var assignment in rule.Descendants().Where(element => string.Equals(element.Name.LocalName, "Action", StringComparison.OrdinalIgnoreCase) || string.Equals(element.Name.LocalName, "Default", StringComparison.OrdinalIgnoreCase)))
                {
                    var target = assignment.Elements().FirstOrDefault(element => string.Equals(element.Name.LocalName, "Target", StringComparison.OrdinalIgnoreCase));
                    if (target is null)
                    {
                        continue;
                    }

                    var categoryName = target.Descendants().FirstOrDefault(element => string.Equals(element.Name.LocalName, "CategoryName", StringComparison.OrdinalIgnoreCase))?.Value?.Trim();
                    var propertyName = target.Descendants().FirstOrDefault(element => string.Equals(element.Name.LocalName, "PropertyName", StringComparison.OrdinalIgnoreCase))?.Value?.Trim();
                    if (!string.Equals(categoryName, "Styles", StringComparison.OrdinalIgnoreCase)
                        || !string.Equals(propertyName, "GraphicalStyle", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var styleName = assignment.Descendants().FirstOrDefault(element => string.Equals(element.Name.LocalName, "Value", StringComparison.OrdinalIgnoreCase))?.Value?.Trim();
                    if (string.IsNullOrWhiteSpace(styleName))
                    {
                        continue;
                    }

                    if (seen.Add(styleName))
                    {
                        orderedStyles.Add(styleName);
                    }
                }

                if (orderedStyles.Count > 0)
                {
                    result[className] = orderedStyles;
                }
            }
        }
        catch
        {
        }

        return result;
    }

    private static IReadOnlyList<string> ResolveStyleCandidates(
        string tableName,
        PaletteCategory category,
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes,
        IReadOnlyDictionary<string, Dictionary<string, PropertyDefinition>> properties,
        IReadOnlyDictionary<string, Dictionary<string, Dictionary<string, string>>> columnAttributes,
        IReadOnlyDictionary<string, IReadOnlyList<string>> xmlStyleCandidates)
    {
        var ordered = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var propertyName in GetStylePropertyNames(category))
        {
            Add(ResolveTableAttributeValue(tableName, tables, attributes, propertyName));
        }

        if (TryResolveClassStringProperty(
                tableName,
                tables,
                attributes,
                properties,
                columnAttributes,
                out var inheritedStyleValue,
                GetStylePropertyNames(category)))
        {
            Add(inheritedStyleValue);
        }

        if (xmlStyleCandidates.TryGetValue(tableName, out var xmlCandidates))
        {
            foreach (var candidate in xmlCandidates)
            {
                Add(candidate);
            }
        }

        return ordered;

        void Add(string? candidate)
        {
            var trimmed = candidate?.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return;
            }

            if (seen.Add(trimmed))
            {
                ordered.Add(trimmed);
            }
        }
    }

    private static string[] GetStylePropertyNames(PaletteCategory category)
    {
        return category == PaletteCategory.Lines
            ? new[] { "GraphicalStyleName", "GraphicalStyle", "LineStyleName", "LineStyle" }
            : new[] { "GraphicalStyleName", "GraphicalStyle", "SymbolStyleName", "SymbolStyle" };
    }

    private static string ResolveDisplayName(
        string tableName,
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes)
    {
        var rawDisplayName = ResolveTableAttributeValue(tableName, tables, attributes, "DisplayName");
        return !string.IsNullOrWhiteSpace(rawDisplayName)
            ? rawDisplayName.Trim()
            : SplitPascalCase(tableName);
    }

    private static string? ResolveSupportedStandardsRaw(
        string tableName,
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes,
        IReadOnlyDictionary<string, Dictionary<string, PropertyDefinition>> properties,
        IReadOnlyDictionary<string, Dictionary<string, Dictionary<string, string>>> columnAttributes)
    {
        var directValue = ResolveTableAttributeValue(tableName, tables, attributes, "SupportedStandard", "SupportedStandards", "BitwiseFlagValue");
        if (!string.IsNullOrWhiteSpace(directValue))
        {
            return directValue;
        }

        return TryResolveClassStringProperty(
                tableName,
                tables,
                attributes,
                properties,
                columnAttributes,
                out var inheritedValue,
                "SupportedStandard",
                "SupportedStandards",
                "BitwiseFlagValue")
            ? inheritedValue
            : null;
    }

    private static string? ResolveTableAttributeValue(
        string tableName,
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes,
        params string[] attributeNames)
    {
        var current = tableName;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (!string.IsNullOrWhiteSpace(current) && visited.Add(current))
        {
            if (attributes.TryGetValue(current, out var tableAttributes))
            {
                foreach (var attributeName in attributeNames)
                {
                    if (tableAttributes.TryGetValue(attributeName, out var attributeValue)
                        && !string.IsNullOrWhiteSpace(attributeValue))
                    {
                        return attributeValue.Trim();
                    }
                }
            }

            current = tables.TryGetValue(current, out var row) ? row.BaseTable ?? string.Empty : string.Empty;
        }

        return null;
    }

    private static bool TryResolveClassStringProperty(
        string tableName,
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes,
        IReadOnlyDictionary<string, Dictionary<string, PropertyDefinition>> properties,
        IReadOnlyDictionary<string, Dictionary<string, Dictionary<string, string>>> columnAttributes,
        out string? value,
        params string[] propertyNames)
    {
        value = null;

        var directValue = ResolveTableAttributeValue(tableName, tables, attributes, propertyNames);
        if (!string.IsNullOrWhiteSpace(directValue))
        {
            value = directValue.Trim();
            return true;
        }

        var current = tableName;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (!string.IsNullOrWhiteSpace(current) && visited.Add(current))
        {
            if (columnAttributes.TryGetValue(current, out var tableColumns))
            {
                foreach (var propertyName in propertyNames)
                {
                    if (tableColumns.TryGetValue(propertyName, out var propertyColumnAttributes))
                    {
                        var columnValue = ResolveColumnStringValue(propertyColumnAttributes);
                        if (!string.IsNullOrWhiteSpace(columnValue))
                        {
                            value = columnValue.Trim();
                            return true;
                        }
                    }
                }
            }

            if (properties.TryGetValue(current, out var tableProperties))
            {
                foreach (var propertyName in propertyNames)
                {
                    if (tableProperties.TryGetValue(propertyName, out var propertyDefinition)
                        && !string.IsNullOrWhiteSpace(propertyDefinition.DefaultValue))
                    {
                        value = propertyDefinition.DefaultValue!.Trim();
                        return true;
                    }
                }
            }

            current = tables.TryGetValue(current, out var row) ? row.BaseTable ?? string.Empty : string.Empty;
        }

        return false;
    }

    private static string? ResolveColumnStringValue(IReadOnlyDictionary<string, string> propertyColumnAttributes)
    {
        foreach (var attributeName in new[] { "VALUE", "Value", "DefaultValue", "DEFAULTVALUE", "AssignedValue", "ASSIGNEDVALUE" })
        {
            if (propertyColumnAttributes.TryGetValue(attributeName, out var value)
                && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static bool TryResolveClassBooleanProperty(
        string tableName,
        string propertyName,
        IReadOnlyDictionary<string, TableRow> tables,
        IReadOnlyDictionary<string, Dictionary<string, string>> attributes,
        IReadOnlyDictionary<string, Dictionary<string, PropertyDefinition>> properties,
        IReadOnlyDictionary<string, Dictionary<string, Dictionary<string, string>>> columnAttributes,
        out bool value)
    {
        value = false;

        var directAttributeValue = ResolveTableAttributeValue(tableName, tables, attributes, propertyName, propertyName.ToUpperInvariant(), SplitPascalCase(propertyName));
        if (TryParseBoolean(directAttributeValue, out value))
        {
            return true;
        }

        var current = tableName;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (!string.IsNullOrWhiteSpace(current) && visited.Add(current))
        {
            if (columnAttributes.TryGetValue(current, out var tableColumns)
                && tableColumns.TryGetValue(propertyName, out var propertyColumnAttributes)
                && TryParseBoolean(ResolveColumnStringValue(propertyColumnAttributes), out value))
            {
                return true;
            }

            if (properties.TryGetValue(current, out var tableProperties)
                && tableProperties.TryGetValue(propertyName, out var propertyDefinition)
                && TryParseBoolean(propertyDefinition.DefaultValue, out value))
            {
                return true;
            }

            current = tables.TryGetValue(current, out var row) ? row.BaseTable ?? string.Empty : string.Empty;
        }

        return false;
    }

    private static bool TryParseBoolean(string? value, out bool result)
    {
        result = false;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "YES", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "1", StringComparison.OrdinalIgnoreCase))
        {
            result = true;
            return true;
        }

        if (string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "NO", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "0", StringComparison.OrdinalIgnoreCase))
        {
            result = false;
            return true;
        }

        return false;
    }

    private static bool IsVisibleInCurrentMode(int? supportedStandardsMask, ProjectRuntimeContext context)
    {
        if (!context.EffectiveRespectSupportedStandards)
        {
            return true;
        }

        return supportedStandardsMask is int mask
            && SupportedStandardsHelper.MatchesProjectStandard(mask, context.CurrentStandardMask);
    }

    private static bool TryResolvePrimaryCategory(
        string tableName,
        IReadOnlyDictionary<string, TableRow> tables,
        out PaletteCategory category)
    {
        if (IsInHierarchy(tableName, "HandValves", tables)
            || string.Equals(tableName, "ControlValve", StringComparison.OrdinalIgnoreCase)
            || IsInHierarchy(tableName, "ReliefValves", tables))
        {
            category = PaletteCategory.Valves;
            return true;
        }

        if (IsInHierarchy(tableName, "PipingFittings", tables))
        {
            category = PaletteCategory.Fittings;
            return true;
        }

        if (IsInHierarchy(tableName, "PipingSpecialtyItems", tables))
        {
            category = PaletteCategory.Speciality;
            return true;
        }

        if (IsInHierarchy(tableName, "Reducers", tables))
        {
            category = PaletteCategory.Reducers;
            return true;
        }

        if (IsInHierarchy(tableName, "Equipment", tables))
        {
            category = PaletteCategory.Equipment;
            return true;
        }

        if (IsInHierarchy(tableName, "Lines", tables))
        {
            category = PaletteCategory.Lines;
            return true;
        }

        if (IsInHierarchy(tableName, "Nozzles", tables))
        {
            category = PaletteCategory.Nozzles;
            return true;
        }

        if (IsInHierarchy(tableName, "NonEngineeringItems", tables))
        {
            category = PaletteCategory.NonEngineeringItems;
            return true;
        }

        if (IsInHierarchy(tableName, "Instrumentation", tables))
        {
            category = PaletteCategory.Instrumentation;
            return true;
        }

        category = default;
        return false;
    }

    private static bool IsInHierarchy(string tableName, string ancestorName, IReadOnlyDictionary<string, TableRow> tables)
    {
        var current = tableName;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (!string.IsNullOrWhiteSpace(current) && visited.Add(current))
        {
            if (string.Equals(current, ancestorName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!tables.TryGetValue(current, out var row))
            {
                return false;
            }

            current = row.BaseTable ?? string.Empty;
        }

        return false;
    }

    private static IReadOnlyList<ProjectPaletteItem> PopulateDescendantItems(ProjectPaletteTreeNode node, bool pruneEmptyChildren)
    {
        for (var i = node.Children.Count - 1; i >= 0; i--)
        {
            var child = node.Children[i];
            PopulateDescendantItems(child, pruneEmptyChildren);
            if (pruneEmptyChildren && child.DescendantItems.Count == 0)
            {
                node.Children.RemoveAt(i);
            }
        }

        node.DescendantItems = CollectDescendantItems(node);
        return node.DescendantItems;
    }

    private static IReadOnlyList<ProjectPaletteItem> CollectDescendantItems(ProjectPaletteTreeNode node)
    {
        var items = new List<ProjectPaletteItem>();
        AddItems(node, items);
        return items
            .GroupBy(item => item.UniqueKey)
            .Select(group => group.First())
            .OrderBy(item => item.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        static void AddItems(ProjectPaletteTreeNode current, ICollection<ProjectPaletteItem> buffer)
        {
            if (current.PaletteItem is not null)
            {
                buffer.Add(current.PaletteItem);
            }

            foreach (var child in current.Children)
            {
                AddItems(child, buffer);
            }
        }
    }

    private static void SortChildren(ProjectPaletteTreeNode node)
    {
        node.Children.Sort(static (left, right) => string.Compare(left.DisplayName, right.DisplayName, StringComparison.OrdinalIgnoreCase));
        foreach (var child in node.Children)
        {
            SortChildren(child);
        }
    }

    private static bool ParseSqliteBoolean(string? value)
    {
        return string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "YES", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "1", StringComparison.OrdinalIgnoreCase);
    }

    private static string SplitPascalCase(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var chars = new List<char>(value.Length * 2);
        for (var i = 0; i < value.Length; i++)
        {
            var ch = value[i];
            if (i > 0 && char.IsUpper(ch) && (char.IsLower(value[i - 1]) || (i + 1 < value.Length && char.IsLower(value[i + 1]))))
            {
                chars.Add(' ');
            }

            chars.Add(ch);
        }

        return new string(chars.ToArray());
    }

    private sealed record TableRow(string TableName, string? BaseTable, bool IsAbstract);

    private sealed record PropertyDefinition(string TableName, string PropertyName, string? PropertyType, string? DefaultValue);

    private sealed record TableCandidate(
        string ClassName,
        string? BaseTable,
        string DisplayName,
        PaletteCategory Category,
        bool IsAbstract,
        IReadOnlyList<string> StyleCandidates,
        int? SupportedStandardsMask,
        bool? TpIncluded);
}
