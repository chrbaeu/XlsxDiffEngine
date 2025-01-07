namespace XlsxDiffTool.Services;

public class PluginService
{
    public IReadOnlyCollection<IPlugin> Plugins { get; } = [];

    public PluginService()
    {
        var pluginTypes = AppDomain.CurrentDomain.GetAssemblies()
            .SelectMany(assembly => assembly.GetTypes())
            .Where(type => typeof(IPlugin).IsAssignableFrom(type) && !type.IsInterface && !type.IsAbstract);
        List<IPlugin> plugins = [];
        foreach (var pluginType in pluginTypes)
        {
            if (Activator.CreateInstance(pluginType) is IPlugin plugin)
            {
                plugins.Add(plugin);
            }
        }
        Plugins = plugins;
    }
}
