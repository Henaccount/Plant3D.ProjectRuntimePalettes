using System.Reflection;
using Plant3D.ProjectRuntimePalettes.Models;
using AcadColors = Autodesk.AutoCAD.Colors;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;

namespace Plant3D.ProjectRuntimePalettes.Services;

public sealed class ProjectStyleLibraryService
{
    private readonly object _syncRoot = new();
    private readonly Dictionary<string, ProjectStyleLibrary> _cache = new(StringComparer.OrdinalIgnoreCase);

    public ProjectStyleResolution Resolve(ProjectPaletteItem item, ProjectRuntimeContext context)
    {
        var library = GetOrLoad(context);
        return Resolve(item, library);
    }

    public ProjectStyleLibrary GetOrLoad(ProjectRuntimeContext context)
    {
        if (string.IsNullOrWhiteSpace(context.SymbolStyleDrawingPath) || !File.Exists(context.SymbolStyleDrawingPath))
        {
            return ProjectStyleLibrary.Empty;
        }

        var cacheKey = $"{context.SymbolStyleDrawingPath}|{File.GetLastWriteTimeUtc(context.SymbolStyleDrawingPath).Ticks}";
        lock (_syncRoot)
        {
            if (_cache.TryGetValue(cacheKey, out var cached))
            {
                return cached;
            }

            var library = Load(context.SymbolStyleDrawingPath);
            _cache[cacheKey] = library;
            return library;
        }
    }

    private static ProjectStyleResolution Resolve(ProjectPaletteItem item, ProjectStyleLibrary library)
    {
        foreach (var candidate in EnumerateLookupKeys(item))
        {
            if (!library.TryResolveStyle(candidate, out var styleInfo))
            {
                continue;
            }

            string? resolvedBlockName = null;
            if (!string.IsNullOrWhiteSpace(styleInfo.SymbolBlockName))
            {
                library.TryResolveBlock(styleInfo.SymbolBlockName, out resolvedBlockName);
            }

            return new ProjectStyleResolution(candidate, styleInfo, styleInfo.SymbolBlockName, resolvedBlockName);
        }

        return new ProjectStyleResolution(null, null, null, null);
    }

    private static IEnumerable<string> EnumerateLookupKeys(ProjectPaletteItem item)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var candidate in item.StyleCandidates)
        {
            if (TryAdd(candidate, seen, out var normalized))
            {
                yield return normalized;
            }
        }

        foreach (var candidate in new[] { item.VisualName, item.SymbolName, item.LineStyleName })
        {
            if (TryAdd(candidate, seen, out var normalized))
            {
                yield return normalized;
            }
        }
    }

    private static bool TryAdd(string? value, ISet<string> seen, out string normalized)
    {
        normalized = string.Empty;
        var trimmed = value?.Trim();
        if (string.IsNullOrWhiteSpace(trimmed) || !seen.Add(trimmed))
        {
            return false;
        }

        normalized = trimmed;
        return true;
    }

    private static ProjectStyleLibrary Load(string symbolStyleDrawingPath)
    {
        var stylesByName = new Dictionary<string, ProjectStyleInfo>(StringComparer.OrdinalIgnoreCase);
        var exactBlocks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var database = new AcadDb.Database(false, true);
            database.ReadDwgFile(symbolStyleDrawingPath, AcadDb.FileOpenMode.OpenForReadAndAllShare, false, null);

            var previousWorkingDatabase = AcadDb.HostApplicationServices.WorkingDatabase;
            try
            {
                AcadDb.HostApplicationServices.WorkingDatabase = database;

                using var transaction = database.TransactionManager.StartTransaction();
                IndexBlocks(database, transaction, exactBlocks);
                LoadGenericStyleMetadata(database, transaction, stylesByName);
                LoadStrictSymbolAssignments(database, transaction, stylesByName, exactBlocks);
                transaction.Commit();
            }
            finally
            {
                try
                {
                    AcadDb.HostApplicationServices.WorkingDatabase = previousWorkingDatabase;
                }
                catch
                {
                }
            }
        }
        catch
        {
        }

        return new ProjectStyleLibrary(stylesByName, exactBlocks);
    }

    private static void IndexBlocks(
        AcadDb.Database database,
        AcadDb.Transaction transaction,
        ISet<string> exactBlocks)
    {
        var blockTable = (AcadDb.BlockTable)transaction.GetObject(database.BlockTableId, AcadDb.OpenMode.ForRead);
        foreach (AcadDb.ObjectId blockId in blockTable)
        {
            AcadDb.BlockTableRecord? blockRecord;
            try
            {
                blockRecord = transaction.GetObject(blockId, AcadDb.OpenMode.ForRead, false) as AcadDb.BlockTableRecord;
            }
            catch
            {
                continue;
            }

            if (blockRecord is null || blockRecord.IsAnonymous || blockRecord.IsLayout)
            {
                continue;
            }

            var blockName = CleanBlockName(blockRecord.Name);
            if (!string.IsNullOrWhiteSpace(blockName))
            {
                exactBlocks.Add(blockName);
            }
        }
    }

    private static void LoadGenericStyleMetadata(
        AcadDb.Database database,
        AcadDb.Transaction transaction,
        IDictionary<string, ProjectStyleInfo> stylesByName)
    {
        var visited = new HashSet<AcadDb.ObjectId>();
        var path = new List<string>();
        VisitObject(database.NamedObjectsDictionaryId, transaction, visited, path, stylesByName);
    }

    private static void VisitObject(
        AcadDb.ObjectId objectId,
        AcadDb.Transaction transaction,
        ISet<AcadDb.ObjectId> visited,
        IList<string> path,
        IDictionary<string, ProjectStyleInfo> stylesByName)
    {
        if (objectId.IsNull || !visited.Add(objectId))
        {
            return;
        }

        AcadDb.DBObject? obj;
        try
        {
            obj = transaction.GetObject(objectId, AcadDb.OpenMode.ForRead, false);
        }
        catch
        {
            return;
        }

        if (obj is null)
        {
            return;
        }

        CaptureStyleFromObject(obj, path, stylesByName);

        if (obj is AcadDb.DBDictionary dictionary)
        {
            foreach (AcadDb.DBDictionaryEntry entry in dictionary)
            {
                path.Add(entry.Key);
                VisitObject(entry.Value, transaction, visited, path, stylesByName);
                path.RemoveAt(path.Count - 1);
            }
        }

        if (!obj.ExtensionDictionary.IsNull)
        {
            path.Add("$XDICT");
            VisitObject(obj.ExtensionDictionary, transaction, visited, path, stylesByName);
            path.RemoveAt(path.Count - 1);
        }
    }

    private static void CaptureStyleFromObject(
        object styleObject,
        IList<string> path,
        IDictionary<string, ProjectStyleInfo> stylesByName)
    {
        if (!LooksLikeStyleCarrier(styleObject, path))
        {
            return;
        }

        var styleName = FirstNonEmptyStringProperty(styleObject, "Name", "StyleName", "KeyName")
            ?? FirstMeaningfulPathSegment(path);

        if (string.IsNullOrWhiteSpace(styleName))
        {
            return;
        }

        var layerName = FirstNonEmptyStringProperty(styleObject, "LayerName", "Layer");
        var linetypeName = FirstNonEmptyStringProperty(styleObject, "LinetypeName", "Linetype", "LineTypeName", "LineType");
        var styleKind = styleObject.GetType().Name;

        var color = TryGetColorInfo(styleObject);
        var lineWeight = TryGetNullableIntProperty(styleObject, "LineWeight", "Weight");

        if (layerName is null
            && linetypeName is null
            && color.argb is null
            && color.index is null
            && lineWeight is null)
        {
            return;
        }

        UpsertStyle(
            stylesByName,
            new ProjectStyleInfo(
                styleName.Trim(),
                null,
                layerName?.Trim(),
                linetypeName?.Trim(),
                color.argb,
                color.index,
                lineWeight,
                styleKind));
    }

    private static void LoadStrictSymbolAssignments(
        AcadDb.Database database,
        AcadDb.Transaction transaction,
        IDictionary<string, ProjectStyleInfo> stylesByName,
        IReadOnlySet<string> exactBlocks)
    {
        AcadDb.DBDictionary? namedObjects;
        try
        {
            namedObjects = transaction.GetObject(database.NamedObjectsDictionaryId, AcadDb.OpenMode.ForRead, false) as AcadDb.DBDictionary;
        }
        catch
        {
            return;
        }

        if (namedObjects is null
            || !TryGetSubDictionary(namedObjects, transaction, "Autodesk_PNP", out var autodeskPnp)
            || !TryGetSubDictionary(autodeskPnp, transaction, "PNP_STYLES", out var styleDictionary))
        {
            return;
        }

        var styleHandlesByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var styleObjectIdsByName = new Dictionary<string, AcadDb.ObjectId>(StringComparer.OrdinalIgnoreCase);

        foreach (AcadDb.DBDictionaryEntry entry in styleDictionary)
        {
            var styleName = entry.Key?.Trim();
            if (string.IsNullOrWhiteSpace(styleName) || entry.Value.IsNull)
            {
                continue;
            }

            styleHandlesByName[styleName] = entry.Value.Handle.ToString();
            styleObjectIdsByName[styleName] = entry.Value;
        }

        var blockNamesByHandle = ExportAssetStyleBlockNamesFromDxf(database);

        foreach (var pair in styleHandlesByName)
        {
            var styleName = pair.Key;
            var handleText = pair.Value;

            blockNamesByHandle.TryGetValue(handleText, out var rawBlockName);
            rawBlockName = CleanBlockName(rawBlockName);

            if (string.IsNullOrWhiteSpace(rawBlockName)
                && styleObjectIdsByName.TryGetValue(styleName, out var objectId))
            {
                rawBlockName = CleanBlockName(TryReadDirectBlockName(transaction, objectId));
            }

            var resolvedBlockName = NormalizeBlockName(rawBlockName, exactBlocks) ?? rawBlockName;

            UpsertStyle(
                stylesByName,
                new ProjectStyleInfo(
                    styleName,
                    resolvedBlockName,
                    null,
                    null,
                    null,
                    null,
                    null,
                    string.IsNullOrWhiteSpace(resolvedBlockName)
                        ? "PNP_STYLES"
                        : $"PNP_STYLES ({handleText})"));
        }
    }

    private static Dictionary<string, string> ExportAssetStyleBlockNamesFromDxf(AcadDb.Database database)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var tempPath = Path.Combine(Path.GetTempPath(), $"Plant3D.ProjectRuntimePalettes.{Guid.NewGuid():N}.dxf");

        try
        {
            database.DxfOut(tempPath, 16, false);

            using var reader = new StreamReader(tempPath);
            string? currentType = null;
            string? currentHandle = null;
            string? currentBlockName = null;

            while (true)
            {
                var codeLine = reader.ReadLine();
                if (codeLine is null)
                {
                    break;
                }

                var valueLine = reader.ReadLine();
                if (valueLine is null)
                {
                    break;
                }

                var code = codeLine.Trim();
                var value = valueLine.Trim();
                if (code.Length == 0)
                {
                    continue;
                }

                if (string.Equals(code, "0", StringComparison.Ordinal))
                {
                    CommitAssetStyle(result, currentType, currentHandle, currentBlockName);
                    currentType = value;
                    currentHandle = null;
                    currentBlockName = null;
                    continue;
                }

                if (!string.Equals(currentType, "ACPPASSETSTYLE", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if ((string.Equals(code, "5", StringComparison.Ordinal) || string.Equals(code, "105", StringComparison.Ordinal)) && string.IsNullOrWhiteSpace(currentHandle))
                {
                    currentHandle = NormalizeHandle(value);
                    continue;
                }

                if (string.Equals(code, "4", StringComparison.Ordinal) && string.IsNullOrWhiteSpace(currentBlockName))
                {
                    currentBlockName = CleanBlockName(value);
                }
            }

            CommitAssetStyle(result, currentType, currentHandle, currentBlockName);
        }
        catch
        {
            result.Clear();
        }
        finally
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
            }
            catch
            {
            }
        }

        return result;
    }

    private static void CommitAssetStyle(
        IDictionary<string, string> result,
        string? currentType,
        string? currentHandle,
        string? currentBlockName)
    {
        if (!string.Equals(currentType, "ACPPASSETSTYLE", StringComparison.OrdinalIgnoreCase)
            || string.IsNullOrWhiteSpace(currentHandle)
            || string.IsNullOrWhiteSpace(currentBlockName)
            || result.ContainsKey(currentHandle))
        {
            return;
        }

        result[currentHandle] = currentBlockName;
    }

    private static string? TryReadDirectBlockName(AcadDb.Transaction transaction, AcadDb.ObjectId objectId)
    {
        try
        {
            if (transaction.GetObject(objectId, AcadDb.OpenMode.ForRead, false) is not AcadDb.DBObject dbObject)
            {
                return null;
            }

            return FirstNonEmptyStringProperty(dbObject, "BlockName", "SymbolBlockName");
        }
        catch
        {
            return null;
        }
    }

    private static bool TryGetSubDictionary(
        AcadDb.DBDictionary parent,
        AcadDb.Transaction transaction,
        string key,
        out AcadDb.DBDictionary dictionary)
    {
        dictionary = null!;

        try
        {
            if (!parent.Contains(key))
            {
                return false;
            }

            dictionary = (AcadDb.DBDictionary)transaction.GetObject(parent.GetAt(key), AcadDb.OpenMode.ForRead);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string? NormalizeBlockName(string? blockName, IReadOnlySet<string> exactBlocks)
    {
        var cleaned = CleanBlockName(blockName);
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            return null;
        }

        return exactBlocks.Contains(cleaned) ? cleaned : null;
    }

    private static string? CleanBlockName(string? blockName)
    {
        var trimmed = blockName?.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return null;
        }

        if ((trimmed.StartsWith("\"", StringComparison.Ordinal) && trimmed.EndsWith("\"", StringComparison.Ordinal))
            || (trimmed.StartsWith("'", StringComparison.Ordinal) && trimmed.EndsWith("'", StringComparison.Ordinal)))
        {
            trimmed = trimmed[1..^1].Trim();
        }

        return string.IsNullOrWhiteSpace(trimmed) ? null : trimmed;
    }

    private static string? NormalizeHandle(string? handleText)
    {
        var trimmed = handleText?.Trim();
        return string.IsNullOrWhiteSpace(trimmed) ? null : trimmed.ToUpperInvariant();
    }

    private static bool LooksLikeStyleCarrier(object instance, IList<string> path)
    {
        var type = instance.GetType();
        var typeName = type.FullName ?? string.Empty;
        if (typeName.StartsWith("Autodesk.ProcessPower", StringComparison.Ordinal)
            || typeName.Contains("Style", StringComparison.Ordinal)
            || typeName.Contains("Symbol", StringComparison.Ordinal))
        {
            return true;
        }

        return path.Any(segment =>
            !string.IsNullOrWhiteSpace(segment)
            && (segment.Contains("Style", StringComparison.OrdinalIgnoreCase)
                || segment.Contains("Symbol", StringComparison.OrdinalIgnoreCase)));
    }

    private static string? FirstNonEmptyStringProperty(object instance, params string[] propertyNames)
    {
        foreach (var propertyName in propertyNames)
        {
            var property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (property is null || property.GetIndexParameters().Length != 0)
            {
                continue;
            }

            string? value;
            try
            {
                value = property.GetValue(instance)?.ToString();
            }
            catch
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static int? TryGetNullableIntProperty(object instance, params string[] propertyNames)
    {
        foreach (var propertyName in propertyNames)
        {
            var property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (property is null || property.GetIndexParameters().Length != 0)
            {
                continue;
            }

            object? value;
            try
            {
                value = property.GetValue(instance);
            }
            catch
            {
                continue;
            }

            switch (value)
            {
                case null:
                    continue;
                case int intValue:
                    return intValue;
                case short shortValue:
                    return shortValue;
            }

            if (int.TryParse(value.ToString(), out var parsed))
            {
                return parsed;
            }
        }

        return null;
    }

    private static string? FirstMeaningfulPathSegment(IList<string> path)
    {
        for (var i = path.Count - 1; i >= 0; i--)
        {
            var segment = path[i]?.Trim();
            if (string.IsNullOrWhiteSpace(segment)
                || string.Equals(segment, "$XDICT", StringComparison.OrdinalIgnoreCase)
                || segment.StartsWith("ACAD_", StringComparison.OrdinalIgnoreCase)
                || string.Equals(segment, "ROOT", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return segment;
        }

        return null;
    }

    private static void UpsertStyle(
        IDictionary<string, ProjectStyleInfo> stylesByName,
        ProjectStyleInfo incoming)
    {
        if (stylesByName.TryGetValue(incoming.StyleName, out var existing))
        {
            stylesByName[incoming.StyleName] = new ProjectStyleInfo(
                incoming.StyleName,
                incoming.SymbolBlockName ?? existing.SymbolBlockName,
                incoming.LayerName ?? existing.LayerName,
                incoming.LinetypeName ?? existing.LinetypeName,
                incoming.ArgbColor ?? existing.ArgbColor,
                incoming.ColorIndex ?? existing.ColorIndex,
                incoming.LineWeight ?? existing.LineWeight,
                string.IsNullOrWhiteSpace(incoming.StyleKind) ? existing.StyleKind : incoming.StyleKind);
            return;
        }

        stylesByName[incoming.StyleName] = incoming;
    }

    private static (int? argb, short? index) TryGetColorInfo(object instance)
    {
        foreach (var propertyName in new[] { "Color", "LineColor", "EntityColor" })
        {
            var property = instance.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (property is null || property.GetIndexParameters().Length != 0)
            {
                continue;
            }

            object? value;
            try
            {
                value = property.GetValue(instance);
            }
            catch
            {
                continue;
            }

            if (value is null)
            {
                continue;
            }

            if (value is AcadColors.Color acadColor)
            {
                try
                {
                    return (acadColor.ColorValue.ToArgb(), acadColor.ColorIndex);
                }
                catch
                {
                    return (MapAciToArgb(acadColor.ColorIndex), acadColor.ColorIndex);
                }
            }

            if (value is short shortIndex)
            {
                return (MapAciToArgb(shortIndex), shortIndex);
            }

            if (value is int intIndex)
            {
                return (MapAciToArgb((short)intIndex), (short)intIndex);
            }

            if (short.TryParse(value.ToString(), out var parsedIndex))
            {
                return (MapAciToArgb(parsedIndex), parsedIndex);
            }
        }

        return (null, null);
    }

    private static int? MapAciToArgb(short aci)
    {
        return aci switch
        {
            1 => System.Drawing.Color.Red.ToArgb(),
            2 => System.Drawing.Color.Yellow.ToArgb(),
            3 => System.Drawing.Color.Lime.ToArgb(),
            4 => System.Drawing.Color.Cyan.ToArgb(),
            5 => System.Drawing.Color.Blue.ToArgb(),
            6 => System.Drawing.Color.Magenta.ToArgb(),
            7 => System.Drawing.Color.Black.ToArgb(),
            8 => System.Drawing.Color.Gray.ToArgb(),
            9 => System.Drawing.Color.LightGray.ToArgb(),
            _ => null
        };
    }
}

public sealed class ProjectStyleLibrary
{
    public static ProjectStyleLibrary Empty { get; } = new(
        new Dictionary<string, ProjectStyleInfo>(StringComparer.OrdinalIgnoreCase),
        new HashSet<string>(StringComparer.OrdinalIgnoreCase));

    public ProjectStyleLibrary(
        IReadOnlyDictionary<string, ProjectStyleInfo> stylesByName,
        IReadOnlySet<string> exactBlocks)
    {
        StylesByName = stylesByName;
        ExactBlocks = exactBlocks;
    }

    public IReadOnlyDictionary<string, ProjectStyleInfo> StylesByName { get; }

    public IReadOnlySet<string> ExactBlocks { get; }

    public bool HasStyleDefinitions => StylesByName.Count > 0;

    public bool TryResolveStyle(string? styleName, out ProjectStyleInfo styleInfo)
    {
        styleInfo = default!;
        if (string.IsNullOrWhiteSpace(styleName))
        {
            return false;
        }

        return StylesByName.TryGetValue(styleName.Trim(), out styleInfo);
    }

    public bool TryResolveBlock(string? blockName, out string? resolvedBlockName)
    {
        resolvedBlockName = null;
        if (string.IsNullOrWhiteSpace(blockName))
        {
            return false;
        }

        var trimmed = blockName.Trim();
        if (!ExactBlocks.Contains(trimmed))
        {
            return false;
        }

        resolvedBlockName = trimmed;
        return true;
    }
}

public sealed record ProjectStyleInfo(
    string StyleName,
    string? SymbolBlockName,
    string? LayerName,
    string? LinetypeName,
    int? ArgbColor,
    short? ColorIndex,
    int? LineWeight,
    string StyleKind)
{
    public bool HasSymbolBlock => !string.IsNullOrWhiteSpace(SymbolBlockName);
}

public sealed record ProjectStyleResolution(
    string? RequestedKey,
    ProjectStyleInfo? StyleInfo,
    string? RawBlockName,
    string? ResolvedBlockName)
{
    public string? DisplayBlockName => ResolvedBlockName ?? RawBlockName;

    public bool HasResolvedBlock => !string.IsNullOrWhiteSpace(ResolvedBlockName);
}
