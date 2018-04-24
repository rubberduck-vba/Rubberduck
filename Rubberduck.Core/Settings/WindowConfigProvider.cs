using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class WindowConfigProvider : IConfigProvider<WindowSettings>
    {
        private readonly IPersistanceService<WindowSettings> _persister;
        private readonly WindowSettings _defaultSettings;

        public WindowConfigProvider(IPersistanceService<WindowSettings> persister)
        {
            _persister = persister;
            _defaultSettings = new DefaultSettings<WindowSettings>().Default;
        }

        public WindowSettings Create()
        {
            return _persister.Load(_defaultSettings) ?? _defaultSettings;
        }

        public WindowSettings CreateDefaults()
        {
            return _defaultSettings;
        }

        public void Save(WindowSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
