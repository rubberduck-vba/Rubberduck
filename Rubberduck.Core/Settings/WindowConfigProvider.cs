using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class WindowConfigProvider : ConfigurationServiceBase<WindowSettings>
    {
        private readonly WindowSettings _defaultSettings;

        public WindowConfigProvider(IPersistanceService<WindowSettings> persister)
            : base(persister)
        {
            _defaultSettings = new DefaultSettings<WindowSettings, Properties.Settings>().Default;
        }

        public override WindowSettings Load()
        {
            return persister.Load(_defaultSettings) ?? _defaultSettings;
        }

        public override WindowSettings LoadDefaults()
        {
            return _defaultSettings;
        }
    }
}
