using Autodesk.AutoCAD.Runtime;
using Plant3D.ProjectRuntimePalettes.Services;

namespace Plant3D.ProjectRuntimePalettes;

public sealed class PluginEntry : IExtensionApplication
{
    public void Initialize()
    {
        PluginDependencyLoader.Initialize();
    }

    public void Terminate()
    {
    }
}
