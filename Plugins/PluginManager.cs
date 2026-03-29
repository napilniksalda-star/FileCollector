using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FileCollector.Plugins
{
    public class PluginManager : IPluginHost
    {
        private List<IPlugin> plugins = new();
        private Action<string>? logAction;

        public PluginManager(Action<string> logAction)
        {
            this.logAction = logAction;
        }

        public void LoadPlugins(string pluginFolder)
        {
            if (!Directory.Exists(pluginFolder))
            {
                Directory.CreateDirectory(pluginFolder);
                Log($"Создана папка плагинов: {pluginFolder}");
                return;
            }

            foreach (var dll in Directory.GetFiles(pluginFolder, "*.dll"))
            {
                try
                {
                    var assembly = Assembly.LoadFrom(dll);
                    var types = assembly.GetTypes()
                        .Where(t => typeof(IPlugin).IsAssignableFrom(t) && !t.IsInterface && !t.IsAbstract);

                    foreach (var type in types)
                    {
                        var plugin = (IPlugin?)Activator.CreateInstance(type);
                        if (plugin != null)
                        {
                            plugin.Initialize(this);
                            plugins.Add(plugin);
                            Log($"Загружен плагин: {plugin.Name} v{plugin.Version}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Ошибка загрузки плагина {dll}: {ex.Message}");
                }
            }
        }

        public T? GetPlugin<T>() where T : IPlugin => plugins.OfType<T>().FirstOrDefault();

        public IEnumerable<T> GetPlugins<T>() where T : IPlugin => plugins.OfType<T>();

        public void Shutdown()
        {
            foreach (var plugin in plugins)
            {
                try
                {
                    plugin.Shutdown();
                }
                catch (Exception ex)
                {
                    Log($"Ошибка при выгрузке плагина {plugin.Name}: {ex.Message}");
                }
            }
            plugins.Clear();
        }

        public void Log(string message) => logAction?.Invoke(message);

        public string GetSetting(string key) => string.Empty;
    }
}
