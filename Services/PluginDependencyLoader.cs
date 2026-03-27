using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;

namespace Plant3D.ProjectRuntimePalettes.Services;

internal static class PluginDependencyLoader
{
    private static readonly object SyncRoot = new();
    private static bool _initialized;
    private static string? _pluginDirectory;
    private static AssemblyDependencyResolver? _resolver;
    private static AssemblyLoadContext? _loadContext;

    public static void Initialize()
    {
        lock (SyncRoot)
        {
            if (_initialized)
            {
                return;
            }

            var pluginAssembly = typeof(PluginEntry).Assembly;
            var pluginAssemblyPath = pluginAssembly.Location;
            if (string.IsNullOrWhiteSpace(pluginAssemblyPath) || !File.Exists(pluginAssemblyPath))
            {
                _initialized = true;
                return;
            }

            _pluginDirectory = Path.GetDirectoryName(pluginAssemblyPath);
            _resolver = new AssemblyDependencyResolver(pluginAssemblyPath);
            _loadContext = AssemblyLoadContext.GetLoadContext(pluginAssembly) ?? AssemblyLoadContext.Default;
            _loadContext.Resolving += OnResolvingManagedAssembly;
            _loadContext.ResolvingUnmanagedDll += OnResolvingUnmanagedDll;
            _initialized = true;
        }
    }

    private static Assembly? OnResolvingManagedAssembly(AssemblyLoadContext loadContext, AssemblyName assemblyName)
    {
        var alreadyLoaded = AppDomain.CurrentDomain
            .GetAssemblies()
            .FirstOrDefault(candidate => AssemblyName.ReferenceMatchesDefinition(candidate.GetName(), assemblyName));

        if (alreadyLoaded is not null)
        {
            return alreadyLoaded;
        }

        var assemblyPath = ResolveManagedAssemblyPath(assemblyName);
        if (assemblyPath is null)
        {
            return null;
        }

        try
        {
            return loadContext.LoadFromAssemblyPath(assemblyPath);
        }
        catch
        {
            return null;
        }
    }

    private static IntPtr OnResolvingUnmanagedDll(Assembly assembly, string unmanagedDllName)
    {
        var nativeLibraryPath = ResolveNativeLibraryPath(unmanagedDllName);
        if (nativeLibraryPath is null)
        {
            return IntPtr.Zero;
        }

        try
        {
            return NativeLibrary.Load(nativeLibraryPath);
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    private static string? ResolveManagedAssemblyPath(AssemblyName assemblyName)
    {
        if (_pluginDirectory is null)
        {
            return null;
        }

        var resolvedByDeps = _resolver?.ResolveAssemblyToPath(assemblyName);
        if (!string.IsNullOrWhiteSpace(resolvedByDeps) && File.Exists(resolvedByDeps))
        {
            return resolvedByDeps;
        }

        var directCandidate = Path.Combine(_pluginDirectory, $"{assemblyName.Name}.dll");
        if (File.Exists(directCandidate))
        {
            return directCandidate;
        }

        return Directory
            .EnumerateFiles(_pluginDirectory, $"{assemblyName.Name}.dll", SearchOption.AllDirectories)
            .FirstOrDefault();
    }

    private static string? ResolveNativeLibraryPath(string unmanagedDllName)
    {
        if (_pluginDirectory is null)
        {
            return null;
        }

        var resolvedByDeps = _resolver?.ResolveUnmanagedDllToPath(unmanagedDllName);
        if (!string.IsNullOrWhiteSpace(resolvedByDeps) && File.Exists(resolvedByDeps))
        {
            return resolvedByDeps;
        }

        foreach (var fileName in GetNativeLibraryFileNameVariants(unmanagedDllName))
        {
            var directCandidate = Path.Combine(_pluginDirectory, fileName);
            if (File.Exists(directCandidate))
            {
                return directCandidate;
            }

            var runtimeCandidate = Path.Combine(_pluginDirectory, "runtimes", "win-x64", "native", fileName);
            if (File.Exists(runtimeCandidate))
            {
                return runtimeCandidate;
            }

            var recursiveCandidate = Directory
                .EnumerateFiles(_pluginDirectory, fileName, SearchOption.AllDirectories)
                .FirstOrDefault();

            if (recursiveCandidate is not null)
            {
                return recursiveCandidate;
            }
        }

        return null;
    }

    private static IEnumerable<string> GetNativeLibraryFileNameVariants(string unmanagedDllName)
    {
        if (string.IsNullOrWhiteSpace(unmanagedDllName))
        {
            yield break;
        }

        if (Path.HasExtension(unmanagedDllName))
        {
            yield return unmanagedDllName;
            yield break;
        }

        yield return $"{unmanagedDllName}.dll";
        yield return unmanagedDllName;
    }
}
