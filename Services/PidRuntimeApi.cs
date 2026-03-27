using System.Reflection;
using Autodesk.ProcessPower.ProjectManager;
using AcadDb = Autodesk.AutoCAD.DatabaseServices;
using AcadGeom = Autodesk.AutoCAD.Geometry;

namespace Plant3D.ProjectRuntimePalettes.Services;

internal sealed class PidRuntimeApi
{
    private readonly Lazy<RuntimeBindings> _bindings;

    public PidRuntimeApi()
    {
        _bindings = new Lazy<RuntimeBindings>(CreateBindings, isThreadSafe: true);
    }

    public bool IsAvailable => _bindings.Value.IsAvailable;

    public string AvailabilityMessage => _bindings.Value.Message ?? string.Empty;

    public bool TryResolveStyleId(string styleName, AcadDb.Database targetDatabase, out AcadDb.ObjectId styleId, out string message)
    {
        styleId = AcadDb.ObjectId.Null;
        message = string.Empty;

        if (string.IsNullOrWhiteSpace(styleName))
        {
            message = "No style name was provided.";
            return false;
        }

        var bindings = _bindings.Value;
        if (bindings.StyleUtilsType is null)
        {
            message = bindings.Message ?? "The Plant P&ID style runtime API is not available.";
            return false;
        }

        var errors = new List<string>();
        var sawOnlyKeyNotFound = false;

        foreach (var method in FindGetStyleIdMethods(bindings.StyleUtilsType))
        {
            foreach (var args in BuildStyleLookupArgumentSets(method, styleName))
            {
                try
                {
                    var result = WithWorkingDatabase(targetDatabase, () => method.Invoke(null, args));
                    if (TryExtractStyleId(method, result, args, out styleId))
                    {
                        message = string.Empty;
                        return true;
                    }
                }
                catch (TargetInvocationException ex)
                {
                    var inner = Unwrap(ex);
                    if (IsKeyNotFound(inner))
                    {
                        sawOnlyKeyNotFound = true;
                        continue;
                    }

                    errors.Add(inner.Message);
                }
                catch (System.Exception ex)
                {
                    if (IsKeyNotFound(ex))
                    {
                        sawOnlyKeyNotFound = true;
                        continue;
                    }

                    errors.Add(ex.Message);
                }
            }
        }

        message = errors.Count > 0
            ? FirstDistinctError(errors)
            : sawOnlyKeyNotFound
                ? $"Style '{styleName}' is not yet available in the current drawing."
                : $"Style '{styleName}' is not available in the current drawing.";
        return false;
    }

    public bool TryEnsureStyleLoaded(string styleName, string? symbolStyleDwgPath, AcadDb.Database targetDatabase, out AcadDb.ObjectId styleId, out string message)
    {
        styleId = AcadDb.ObjectId.Null;
        if (TryResolveStyleId(styleName, targetDatabase, out styleId, out message))
        {
            return true;
        }

        var bindings = _bindings.Value;
        if (bindings.ProjectSymbolStyleUtilsType is null)
        {
            message = bindings.Message ?? "The Plant project style-copy API is not available.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(symbolStyleDwgPath) || !File.Exists(symbolStyleDwgPath))
        {
            message = $"Style '{styleName}' is not present in the current drawing, and projSymbolStyle.dwg could not be found.";
            return false;
        }

        var copyMethods = FindCopyStyleMethods(bindings.ProjectSymbolStyleUtilsType).ToList();
        if (copyMethods.Count == 0)
        {
            message = $"Could not locate {bindings.ProjectSymbolStyleUtilsType.FullName}.CopyStyle(...).";
            return false;
        }

        var errors = new List<string>();
        var sawOnlyKeyNotFound = false;

        using var sourceDatabase = new AcadDb.Database(false, true);
        sourceDatabase.ReadDwgFile(symbolStyleDwgPath, AcadDb.FileOpenMode.OpenForReadAndAllShare, false, null);

        foreach (var method in copyMethods)
        {
            foreach (var plan in BuildCopyStylePlans(method, styleName, sourceDatabase, targetDatabase))
            {
                try
                {
                    WithWorkingDatabase(plan.WorkingDatabase, () => method.Invoke(null, plan.Arguments));

                    if (TryResolveStyleId(styleName, targetDatabase, out styleId, out _))
                    {
                        message = string.Empty;
                        return true;
                    }
                }
                catch (TargetInvocationException ex)
                {
                    var inner = Unwrap(ex);
                    if (IsKeyNotFound(inner))
                    {
                        sawOnlyKeyNotFound = true;
                        continue;
                    }

                    errors.Add(inner.Message);
                }
                catch (System.Exception ex)
                {
                    if (IsKeyNotFound(ex))
                    {
                        sawOnlyKeyNotFound = true;
                        continue;
                    }

                    errors.Add(ex.Message);
                }
            }
        }

        message = errors.Count > 0
            ? $"Style '{styleName}' could not be copied into the current drawing. {FirstDistinctError(errors)}"
            : sawOnlyKeyNotFound
                ? $"Style '{styleName}' could not be copied into the current drawing. The Plant style API reported eKeyNotFound for all CopyStyle variants."
                : $"Style '{styleName}' could not be copied from projSymbolStyle.dwg into the current drawing.";
        return false;
    }

    public bool TryInsertAsset(string className, string styleName, AcadDb.ObjectId styleId, AcadGeom.Point3d position, AcadDb.Database targetDatabase, out string message)
    {
        var bindings = _bindings.Value;
        if (bindings.AssetAdderType is null)
        {
            message = bindings.Message ?? "The Plant P&ID asset insertion API is not available.";
            return false;
        }

        var attemptMessages = new List<string>();

        foreach (var adder in CreateAdderCandidates(bindings.AssetAdderType, targetDatabase))
        {
            try
            {
                if (TryInvokeAssetPointOverloads(bindings, adder, position, out var pointMessage, styleName, className))
                {
                    message = string.Empty;
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(pointMessage))
                {
                    attemptMessages.Add(pointMessage);
                }

                if (TryAddConfiguredAsset(bindings, adder, className, styleId, position, out var configuredMessage))
                {
                    message = string.Empty;
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(configuredMessage))
                {
                    attemptMessages.Add(configuredMessage);
                }
            }
            finally
            {
                DisposeIfNeeded(adder);
            }
        }

        message = attemptMessages.Count > 0
            ? FirstDistinctError(attemptMessages)
            : $"No compatible asset insertion overload was found for class '{className}'.";
        return false;
    }

    public bool TryInsertLine(string className, string styleName, AcadDb.ObjectId styleId, AcadGeom.Point3dCollection vertices, AcadDb.Database targetDatabase, out string message)
    {
        var bindings = _bindings.Value;
        if (bindings.LineSegmentAdderType is null)
        {
            message = bindings.Message ?? "The Plant P&ID line insertion API is not available.";
            return false;
        }

        var attemptMessages = new List<string>();

        foreach (var adder in CreateAdderCandidates(bindings.LineSegmentAdderType, targetDatabase))
        {
            try
            {
                if (TryInvokeLinePointCollectionOverloads(bindings, adder, vertices, out var pointMessage, styleName, className))
                {
                    message = string.Empty;
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(pointMessage))
                {
                    attemptMessages.Add(pointMessage);
                }

                if (TryAddConfiguredLine(bindings, adder, className, styleId, vertices, out var configuredMessage))
                {
                    message = string.Empty;
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(configuredMessage))
                {
                    attemptMessages.Add(configuredMessage);
                }
            }
            finally
            {
                DisposeIfNeeded(adder);
            }
        }

        message = attemptMessages.Count > 0
            ? FirstDistinctError(attemptMessages)
            : $"No compatible line insertion overload was found for class '{className}'.";
        return false;
    }

    private static RuntimeBindings CreateBindings()
    {
        var assemblies = LoadCandidateAssemblies();

        var assetType = FindType(assemblies, "Asset", "PnIDObjects");
        var dynamicAssetType = FindType(assemblies, "DynamicAsset", "PnIDObjects");
        var assetStyleType = FindType(assemblies, "AssetStyle", "PnIDObjects");
        var assetUtilType = FindType(assemblies, "AssetUtil", "PnIDObjects");
        var assetAdderType = FindType(assemblies, "AssetAdder", "PnIDObjects");
        var lineSegmentType = FindType(assemblies, "LineSegment", "PnIDObjects");
        var lineSegmentAdderType = FindType(assemblies, "LineSegmentAdder", "PnIDObjects");
        var styleUtilsType = FindType(assemblies, "StyleUtils", "ProcessPower");
        var projectSymbolStyleUtilsType = FindType(assemblies, "ProjectSymbolStyleUtils", "ProcessPower");

        var missing = new List<string>();
        if (assetType is null) missing.Add("Asset");
        if (assetAdderType is null) missing.Add("AssetAdder");
        if (lineSegmentType is null) missing.Add("LineSegment");
        if (lineSegmentAdderType is null) missing.Add("LineSegmentAdder");
        if (styleUtilsType is null) missing.Add("StyleUtils");
        if (projectSymbolStyleUtilsType is null) missing.Add("ProjectSymbolStyleUtils");

        var message = missing.Count == 0
            ? string.Empty
            : "The following Plant P&ID runtime types could not be located: " + string.Join(", ", missing) + ".";

        return new RuntimeBindings(
            assetType,
            dynamicAssetType,
            assetStyleType,
            assetUtilType,
            assetAdderType,
            lineSegmentType,
            lineSegmentAdderType,
            styleUtilsType,
            projectSymbolStyleUtilsType,
            message);
    }

    private static IReadOnlyList<Assembly> LoadCandidateAssemblies()
    {
        var loaded = AppDomain.CurrentDomain.GetAssemblies()
            .Where(assembly => !assembly.IsDynamic)
            .ToDictionary(assembly => assembly.GetName().Name ?? string.Empty, assembly => assembly, StringComparer.OrdinalIgnoreCase);

        var root = Path.GetDirectoryName(typeof(Project).Assembly.Location);
        if (!string.IsNullOrWhiteSpace(root) && Directory.Exists(root))
        {
            foreach (var filePath in Directory.EnumerateFiles(root, "*.dll", SearchOption.TopDirectoryOnly))
            {
                var fileName = Path.GetFileName(filePath);
                if (!LooksLikeRelevantPlantAssembly(fileName))
                {
                    continue;
                }

                AssemblyName assemblyName;
                try
                {
                    assemblyName = AssemblyName.GetAssemblyName(filePath);
                }
                catch
                {
                    continue;
                }

                var simpleName = assemblyName.Name ?? string.Empty;
                if (loaded.ContainsKey(simpleName))
                {
                    continue;
                }

                try
                {
                    loaded[simpleName] = Assembly.LoadFrom(filePath);
                }
                catch
                {
                }
            }
        }

        return loaded.Values.ToList();
    }

    private static bool LooksLikeRelevantPlantAssembly(string fileName)
    {
        return fileName.Contains("Pn", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("Plant", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("ProcessPower", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("DataLinks", StringComparison.OrdinalIgnoreCase)
            || fileName.Contains("Style", StringComparison.OrdinalIgnoreCase);
    }

    private static Type? FindType(IEnumerable<Assembly> assemblies, string simpleName, string namespaceHint)
    {
        foreach (var assembly in assemblies)
        {
            foreach (var type in GetTypesSafely(assembly))
            {
                if (!string.Equals(type.Name, simpleName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var ns = type.Namespace ?? string.Empty;
                if (ns.Contains(namespaceHint, StringComparison.OrdinalIgnoreCase)
                    || ns.Contains("ProcessPower", StringComparison.OrdinalIgnoreCase))
                {
                    return type;
                }
            }
        }

        return null;
    }

    private static IEnumerable<Type> GetTypesSafely(Assembly assembly)
    {
        try
        {
            return assembly.GetTypes();
        }
        catch (ReflectionTypeLoadException ex)
        {
            return ex.Types.OfType<Type>();
        }
        catch
        {
            return Array.Empty<Type>();
        }
    }

    private static IEnumerable<MethodInfo> FindGetStyleIdMethods(Type styleUtilsType)
    {
        return styleUtilsType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => string.Equals(method.Name, "GetStyleIdFromName", StringComparison.OrdinalIgnoreCase))
            .Where(method => method.GetParameters().Length >= 1 && method.GetParameters()[0].ParameterType == typeof(string))
            .OrderBy(method => method.GetParameters().Length)
            .ToList();
    }

    private static IEnumerable<object?[]> BuildStyleLookupArgumentSets(MethodInfo method, string styleName)
    {
        var parameters = method.GetParameters();
        if (parameters.Length == 0 || parameters[0].ParameterType != typeof(string))
        {
            yield break;
        }

        var boolValueSets = BuildBoolPermutations(parameters.Skip(1).Where(IsPureBoolParameter).Count()).ToList();
        if (boolValueSets.Count == 0)
        {
            boolValueSets.Add(Array.Empty<bool>());
        }

        foreach (var boolValues in boolValueSets)
        {
            var args = new object?[parameters.Length];
            args[0] = styleName;
            var boolIndex = 0;
            var supported = true;

            for (var index = 1; index < parameters.Length; index++)
            {
                var parameter = parameters[index];
                var parameterType = parameter.ParameterType;
                var elementType = parameterType.IsByRef ? parameterType.GetElementType() : parameterType;

                if (elementType == typeof(bool))
                {
                    args[index] = boolValues[boolIndex++];
                    continue;
                }

                if (parameterType.IsByRef && elementType == typeof(AcadDb.ObjectId))
                {
                    args[index] = AcadDb.ObjectId.Null;
                    continue;
                }

                supported = false;
                break;
            }

            if (supported)
            {
                yield return args;
            }
        }
    }

    private static bool TryExtractStyleId(MethodInfo method, object? result, object?[] args, out AcadDb.ObjectId styleId)
    {
        styleId = AcadDb.ObjectId.Null;

        if (result is AcadDb.ObjectId directId)
        {
            styleId = directId;
            return !styleId.IsNull;
        }

        var parameters = method.GetParameters();
        for (var index = 0; index < parameters.Length; index++)
        {
            var parameter = parameters[index];
            var parameterType = parameter.ParameterType;
            if (!parameterType.IsByRef || parameterType.GetElementType() != typeof(AcadDb.ObjectId))
            {
                continue;
            }

            if (args[index] is AcadDb.ObjectId byRefId)
            {
                styleId = byRefId;
                return !styleId.IsNull && (result is not bool boolResult || boolResult);
            }
        }

        return false;
    }

    private static IEnumerable<MethodInfo> FindCopyStyleMethods(Type projectSymbolStyleUtilsType)
    {
        return projectSymbolStyleUtilsType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => string.Equals(method.Name, "CopyStyle", StringComparison.OrdinalIgnoreCase))
            .Where(method => method.GetParameters().Any(parameter => parameter.ParameterType == typeof(string)))
            .OrderBy(method => method.GetParameters().Length)
            .ToList();
    }

    private static IEnumerable<CopyStylePlan> BuildCopyStylePlans(MethodInfo method, string styleName, AcadDb.Database sourceDatabase, AcadDb.Database targetDatabase)
    {
        var parameters = method.GetParameters();
        var boolCount = parameters.Count(IsPureBoolParameter);
        var boolValueSets = BuildBoolPermutations(boolCount).ToList();
        if (boolValueSets.Count == 0)
        {
            boolValueSets.Add(Array.Empty<bool>());
        }

        foreach (var boolValues in boolValueSets)
        {
            foreach (var reverseDatabases in new[] { false, true })
            {
                var firstDatabase = reverseDatabases ? targetDatabase : sourceDatabase;
                var secondDatabase = reverseDatabases ? sourceDatabase : targetDatabase;

                foreach (var workingDatabase in new[] { targetDatabase, sourceDatabase })
                {
                    var args = new object?[parameters.Length];
                    var boolIndex = 0;
                    var databaseIndex = 0;
                    var supported = true;

                    for (var index = 0; index < parameters.Length; index++)
                    {
                        var parameter = parameters[index];
                        if (parameter.ParameterType == typeof(string))
                        {
                            args[index] = styleName;
                            continue;
                        }

                        if (IsPureBoolParameter(parameter))
                        {
                            args[index] = boolValues[boolIndex++];
                            continue;
                        }

                        if (parameter.ParameterType == typeof(AcadDb.Database))
                        {
                            args[index] = databaseIndex++ == 0 ? firstDatabase : secondDatabase;
                            continue;
                        }

                        supported = false;
                        break;
                    }

                    if (supported && databaseIndex == 2)
                    {
                        yield return new CopyStylePlan(args, workingDatabase);
                    }
                }
            }
        }
    }

    private static IEnumerable<bool[]> BuildBoolPermutations(int count)
    {
        if (count <= 0)
        {
            yield break;
        }

        var max = 1 << count;
        for (var mask = 0; mask < max; mask++)
        {
            var values = new bool[count];
            for (var index = 0; index < count; index++)
            {
                values[index] = (mask & (1 << index)) != 0;
            }

            yield return values;
        }
    }

    private static bool IsPureBoolParameter(ParameterInfo parameter)
    {
        return !parameter.ParameterType.IsByRef && parameter.ParameterType == typeof(bool);
    }

    private static IEnumerable<object> CreateAdderCandidates(Type adderType, AcadDb.Database targetDatabase)
    {
        foreach (var constructor in adderType.GetConstructors(BindingFlags.Public | BindingFlags.Instance)
                     .OrderByDescending(ctor => ctor.GetParameters().Length))
        {
            var parameters = constructor.GetParameters();
            if (parameters.Length == 2
                && parameters[0].ParameterType == typeof(AcadDb.Database)
                && parameters[1].ParameterType == typeof(bool))
            {
                foreach (var boolValue in new[] { true, false })
                {
                    object? instance;
                    try
                    {
                        instance = constructor.Invoke(new object?[] { targetDatabase, boolValue });
                    }
                    catch
                    {
                        continue;
                    }

                    if (instance is not null)
                    {
                        yield return instance;
                    }
                }
            }

            if (parameters.Length == 1 && parameters[0].ParameterType == typeof(AcadDb.Database))
            {
                object? instance;
                try
                {
                    instance = constructor.Invoke(new object?[] { targetDatabase });
                }
                catch
                {
                    continue;
                }

                if (instance is not null)
                {
                    yield return instance;
                }
            }
        }
    }

    private static bool TryAddConfiguredAsset(RuntimeBindings bindings, object assetAdder, string className, AcadDb.ObjectId styleId, AcadGeom.Point3d position, out string message)
    {
        message = string.Empty;

        object? asset = null;
        try
        {
            var assetType = ResolveAssetCreationType(bindings, styleId);
            asset = Activator.CreateInstance(assetType);
            if (asset is null)
            {
                message = $"Could not create an instance of {assetType.FullName}.";
                return false;
            }

            TryInvokeMethod(asset, "SetDatabaseDefaults");
            TrySetProperty(asset, "ClassName", className);
            TrySetProperty(asset, "StyleId", styleId);
            TrySetProperty(asset, "StyleID", styleId);
            TrySetProperty(asset, "Position", position);
            TrySetProperty(asset, "Normal", AcadGeom.Vector3d.ZAxis);
            TrySetProperty(asset, "XAxis", AcadGeom.Vector3d.XAxis);
            TrySetProperty(asset, "ScaleFactors", new AcadGeom.Scale3d(1.0));

            TryInvokeMethod(asset, "Initialize");
            TryInvokeMethod(asset, "ReApplyStyle");
            TryInvokeMethod(asset, "Refresh");

            var addMethod = bindings.AssetAdderType!
                .GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(method => string.Equals(method.Name, "Add", StringComparison.OrdinalIgnoreCase))
                .Where(method =>
                {
                    var parameters = method.GetParameters();
                    return parameters.Length == 1
                        && parameters[0].ParameterType.IsAssignableFrom(asset.GetType());
                })
                .OrderBy(method => method.GetParameters()[0].ParameterType == asset.GetType() ? 0 : 1)
                .FirstOrDefault();

            if (addMethod is null)
            {
                message = $"No single-parameter {bindings.AssetAdderType.FullName}.Add({asset.GetType().FullName}) overload is available.";
                return false;
            }

            addMethod.Invoke(assetAdder, new object?[] { asset });
            return true;
        }
        catch (TargetInvocationException ex)
        {
            message = $"Configured asset insertion failed: {Unwrap(ex).Message}";
            return false;
        }
        catch (System.Exception ex)
        {
            message = $"Configured asset insertion failed: {ex.Message}";
            return false;
        }
        finally
        {
            DisposeIfNeeded(asset);
        }
    }

    private static Type ResolveAssetCreationType(RuntimeBindings bindings, AcadDb.ObjectId styleId)
    {
        if (RequiresDynamicAsset(bindings, styleId) && bindings.DynamicAssetType is not null)
        {
            return bindings.DynamicAssetType;
        }

        return bindings.AssetType!;
    }

    private static bool RequiresDynamicAsset(RuntimeBindings bindings, AcadDb.ObjectId styleId)
    {
        try
        {
            if (bindings.AssetStyleType is not null)
            {
                var styleMethod = bindings.AssetStyleType
                    .GetMethods(BindingFlags.Public | BindingFlags.Static)
                    .FirstOrDefault(method => string.Equals(method.Name, "StyleRequiresDynamicAsset", StringComparison.OrdinalIgnoreCase)
                        && method.ReturnType == typeof(bool)
                        && method.GetParameters().Length == 1
                        && method.GetParameters()[0].ParameterType == typeof(AcadDb.ObjectId));

                if (styleMethod?.Invoke(null, new object?[] { styleId }) is bool styleRequires)
                {
                    return styleRequires;
                }
            }
        }
        catch
        {
        }

        try
        {
            if (bindings.AssetUtilType is not null)
            {
                var utilMethod = bindings.AssetUtilType
                    .GetMethods(BindingFlags.Public | BindingFlags.Static)
                    .FirstOrDefault(method => string.Equals(method.Name, "IsDynamicAsset", StringComparison.OrdinalIgnoreCase)
                        && method.ReturnType == typeof(bool)
                        && method.GetParameters().Length == 1
                        && method.GetParameters()[0].ParameterType == typeof(AcadDb.ObjectId));

                if (utilMethod?.Invoke(null, new object?[] { styleId }) is bool utilRequires)
                {
                    return utilRequires;
                }
            }
        }
        catch
        {
        }

        return false;
    }

    private static bool TryInvokeAssetPointOverloads(RuntimeBindings bindings, object assetAdder, AcadGeom.Point3d position, out string message, params string[] candidates)
    {
        message = string.Empty;
        var methods = bindings.AssetAdderType!
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(method => string.Equals(method.Name, "Add", StringComparison.OrdinalIgnoreCase))
            .OrderBy(GetAssetPointOverloadPriority)
            .ToList();

        var errorMessages = new List<string>();
        foreach (var stringValue in EnumerateDistinctStringCandidates(candidates))
        {
            foreach (var method in methods)
            {
                if (!TryBuildAssetPointArguments(method, position, stringValue, out var args))
                {
                    continue;
                }

                try
                {
                    method.Invoke(assetAdder, args);
                    return true;
                }
                catch (TargetInvocationException ex)
                {
                    errorMessages.Add($"{method.Name}({stringValue}) failed: {Unwrap(ex).Message}");
                }
                catch (System.Exception ex)
                {
                    errorMessages.Add($"{method.Name}({stringValue}) failed: {ex.Message}");
                }
            }
        }

        message = errorMessages.Count > 0
            ? FirstDistinctError(errorMessages)
            : $"No compatible {bindings.AssetAdderType.FullName}.Add(Point3d, ...) overload was found.";
        return false;
    }

    private static bool TryBuildAssetPointArguments(MethodInfo method, AcadGeom.Point3d position, string stringValue, out object?[] args)
    {
        args = Array.Empty<object?>();
        var parameters = method.GetParameters();

        if (parameters.Length == 2
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(string))
        {
            args = new object?[] { position, stringValue };
            return true;
        }

        if (parameters.Length == 3
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(string)
            && parameters[2].ParameterType == typeof(AcadGeom.Scale3d))
        {
            args = new object?[] { position, stringValue, new AcadGeom.Scale3d(1.0) };
            return true;
        }

        if (parameters.Length == 3
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(double)
            && parameters[2].ParameterType == typeof(string))
        {
            args = new object?[] { position, 0d, stringValue };
            return true;
        }

        if (parameters.Length == 4
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(double)
            && parameters[2].ParameterType == typeof(string)
            && parameters[3].ParameterType == typeof(AcadGeom.Scale3d))
        {
            args = new object?[] { position, 0d, stringValue, new AcadGeom.Scale3d(1.0) };
            return true;
        }

        return false;
    }

    private static bool TryAddConfiguredLine(RuntimeBindings bindings, object adder, string className, AcadDb.ObjectId styleId, AcadGeom.Point3dCollection vertices, out string message)
    {
        message = string.Empty;

        object? lineSegment = null;
        try
        {
            lineSegment = Activator.CreateInstance(bindings.LineSegmentType!);
            if (lineSegment is null)
            {
                message = $"Could not create an instance of {bindings.LineSegmentType!.FullName}.";
                return false;
            }

            TrySetProperty(lineSegment, "ClassName", className);
            TrySetProperty(lineSegment, "StyleId", styleId);
            TrySetProperty(lineSegment, "StyleID", styleId);

            if (!TrySetProperty(lineSegment, "Vertices", vertices))
            {
                TryPopulateVerticesViaMethod(lineSegment, vertices);
            }

            var addMethod = bindings.LineSegmentAdderType!
                .GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(method => string.Equals(method.Name, "Add", StringComparison.OrdinalIgnoreCase))
                .Where(method =>
                {
                    var parameters = method.GetParameters();
                    return parameters.Length == 1
                        && parameters[0].ParameterType.IsAssignableFrom(lineSegment.GetType());
                })
                .OrderBy(method => method.GetParameters()[0].ParameterType == lineSegment.GetType() ? 0 : 1)
                .FirstOrDefault();

            if (addMethod is null)
            {
                message = $"No single-parameter {bindings.LineSegmentAdderType.FullName}.Add({bindings.LineSegmentType.FullName}) overload is available.";
                return false;
            }

            addMethod.Invoke(adder, new object?[] { lineSegment });
            return true;
        }
        catch (TargetInvocationException ex)
        {
            message = $"Configured line insertion failed: {Unwrap(ex).Message}";
            return false;
        }
        catch (System.Exception ex)
        {
            message = $"Configured line insertion failed: {ex.Message}";
            return false;
        }
        finally
        {
            DisposeIfNeeded(lineSegment);
        }
    }

    private static bool TryInvokeLinePointCollectionOverloads(RuntimeBindings bindings, object adder, AcadGeom.Point3dCollection vertices, out string message, params string[] candidates)
    {
        message = string.Empty;
        var methods = bindings.LineSegmentAdderType!
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(method => string.Equals(method.Name, "Add", StringComparison.OrdinalIgnoreCase))
            .OrderBy(GetLinePointOverloadPriority)
            .ToList();

        var errorMessages = new List<string>();

        foreach (var stringValue in EnumerateDistinctStringCandidates(candidates))
        {
            foreach (var method in methods)
            {
                var parameters = method.GetParameters();
                if (parameters.Length == 2
                    && parameters[0].ParameterType == typeof(AcadGeom.Point3dCollection)
                    && parameters[1].ParameterType == typeof(string))
                {
                    try
                    {
                        method.Invoke(adder, new object?[] { vertices, stringValue });
                        return true;
                    }
                    catch (TargetInvocationException ex)
                    {
                        errorMessages.Add($"{method.Name}({stringValue}) failed: {Unwrap(ex).Message}");
                    }
                    catch (System.Exception ex)
                    {
                        errorMessages.Add($"{method.Name}({stringValue}) failed: {ex.Message}");
                    }
                }
            }
        }

        foreach (var method in methods)
        {
            var parameters = method.GetParameters();
            if (parameters.Length == 1 && parameters[0].ParameterType == typeof(AcadGeom.Point3dCollection))
            {
                try
                {
                    method.Invoke(adder, new object?[] { vertices });
                    return true;
                }
                catch (TargetInvocationException ex)
                {
                    errorMessages.Add($"{method.Name}(Point3dCollection) failed: {Unwrap(ex).Message}");
                }
                catch (System.Exception ex)
                {
                    errorMessages.Add($"{method.Name}(Point3dCollection) failed: {ex.Message}");
                }
            }
        }

        message = errorMessages.Count > 0
            ? FirstDistinctError(errorMessages)
            : $"No safe {bindings.LineSegmentAdderType.FullName}.Add(Point3dCollection, ...) overload was found.";
        return false;
    }

    private static int GetAssetPointOverloadPriority(MethodInfo method)
    {
        var parameters = method.GetParameters();
        if (parameters.Length == 2
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(string))
        {
            return 0;
        }

        if (parameters.Length == 3
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(string)
            && parameters[2].ParameterType == typeof(AcadGeom.Scale3d))
        {
            return 1;
        }

        if (parameters.Length == 3
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(double)
            && parameters[2].ParameterType == typeof(string))
        {
            return 2;
        }

        if (parameters.Length == 4
            && parameters[0].ParameterType == typeof(AcadGeom.Point3d)
            && parameters[1].ParameterType == typeof(double)
            && parameters[2].ParameterType == typeof(string)
            && parameters[3].ParameterType == typeof(AcadGeom.Scale3d))
        {
            return 3;
        }

        return int.MaxValue;
    }

    private static int GetLinePointOverloadPriority(MethodInfo method)
    {
        var parameters = method.GetParameters();
        if (parameters.Length == 2
            && parameters[0].ParameterType == typeof(AcadGeom.Point3dCollection)
            && parameters[1].ParameterType == typeof(string))
        {
            return 0;
        }

        if (parameters.Length == 1
            && parameters[0].ParameterType == typeof(AcadGeom.Point3dCollection))
        {
            return 1;
        }

        return int.MaxValue;
    }

    private static void TryPopulateVerticesViaMethod(object lineSegment, AcadGeom.Point3dCollection vertices)
    {
        var addVertexAt = lineSegment.GetType().GetMethod(
            "AddVertexAt",
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
            binder: null,
            types: new[] { typeof(int), typeof(AcadGeom.Point3d) },
            modifiers: null);

        if (addVertexAt is null)
        {
            return;
        }

        for (var index = 0; index < vertices.Count; index++)
        {
            addVertexAt.Invoke(lineSegment, new object?[] { index, vertices[index] });
        }
    }

    private static bool TrySetProperty(object instance, string propertyName, object? value)
    {
        var property = instance.GetType().GetProperty(
            propertyName,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        if (property is null || !property.CanWrite)
        {
            return false;
        }

        try
        {
            property.SetValue(instance, value);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void TryInvokeMethod(object instance, string methodName)
    {
        var method = instance.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase, binder: null, types: Type.EmptyTypes, modifiers: null);
        method?.Invoke(instance, Array.Empty<object?>());
    }

    private static IEnumerable<string> EnumerateDistinctStringCandidates(params string[] values)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values)
        {
            var trimmed = value?.Trim();
            if (string.IsNullOrWhiteSpace(trimmed) || !seen.Add(trimmed))
            {
                continue;
            }

            yield return trimmed;
        }
    }

    private static T WithWorkingDatabase<T>(AcadDb.Database workingDatabase, Func<T> action)
    {
        var previousWorkingDatabase = AcadDb.HostApplicationServices.WorkingDatabase;
        try
        {
            AcadDb.HostApplicationServices.WorkingDatabase = workingDatabase;
            return action();
        }
        finally
        {
            TryRestoreWorkingDatabase(previousWorkingDatabase);
        }
    }

    private static void WithWorkingDatabase(AcadDb.Database workingDatabase, Action action)
    {
        var previousWorkingDatabase = AcadDb.HostApplicationServices.WorkingDatabase;
        try
        {
            AcadDb.HostApplicationServices.WorkingDatabase = workingDatabase;
            action();
        }
        finally
        {
            TryRestoreWorkingDatabase(previousWorkingDatabase);
        }
    }

    private static void TryRestoreWorkingDatabase(AcadDb.Database previousWorkingDatabase)
    {
        try
        {
            AcadDb.HostApplicationServices.WorkingDatabase = previousWorkingDatabase;
        }
        catch
        {
        }
    }

    private static bool IsKeyNotFound(System.Exception ex)
    {
        return ex is KeyNotFoundException
            || ex.Message.Contains("eKeyNotFound", StringComparison.OrdinalIgnoreCase);
    }

    private static string FirstDistinctError(IEnumerable<string> errors)
    {
        return errors
            .Select(error => error?.Trim())
            .Where(error => !string.IsNullOrWhiteSpace(error))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault()
            ?? string.Empty;
    }

    private static void DisposeIfNeeded(object? instance)
    {
        if (instance is IDisposable disposable)
        {
            disposable.Dispose();
        }
    }

    private static System.Exception Unwrap(TargetInvocationException ex)
    {
        return ex.InnerException ?? ex;
    }

    private sealed record CopyStylePlan(object?[] Arguments, AcadDb.Database WorkingDatabase);

    private sealed record RuntimeBindings(
        Type? AssetType,
        Type? DynamicAssetType,
        Type? AssetStyleType,
        Type? AssetUtilType,
        Type? AssetAdderType,
        Type? LineSegmentType,
        Type? LineSegmentAdderType,
        Type? StyleUtilsType,
        Type? ProjectSymbolStyleUtilsType,
        string? Message)
    {
        public bool IsAvailable => AssetType is not null
            && AssetAdderType is not null
            && LineSegmentType is not null
            && LineSegmentAdderType is not null
            && StyleUtilsType is not null
            && ProjectSymbolStyleUtilsType is not null;
    }
}
