using System;
using Mdr.Revit.Infra.Config;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class SettingsCommand
    {
        private readonly ConfigLoader _configLoader;

        public SettingsCommand()
            : this(new ConfigLoader())
        {
        }

        internal SettingsCommand(ConfigLoader configLoader)
        {
            _configLoader = configLoader ?? throw new ArgumentNullException(nameof(configLoader));
        }

        public string Id => "mdr.settings";

        public PluginConfig Load(string configPath)
        {
            return _configLoader.Load(configPath);
        }

        public void Save(string configPath, PluginConfig config)
        {
            _configLoader.Save(configPath, config);
        }
    }
}
