using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class WindowConfigProvider : IConfigProvider<WindowSettings>
    {
        private readonly IPersistanceService<WindowSettings> _persister;

        public WindowConfigProvider(IPersistanceService<WindowSettings> persister)
        {
            _persister = persister;
        }

        public WindowSettings Create()
        {
            var prototype = new WindowSettings();
            return _persister.Load(prototype) ?? prototype;
        }

        public WindowSettings CreateDefaults()
        {
            return new WindowSettings();
        }

        public void Save(WindowSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
