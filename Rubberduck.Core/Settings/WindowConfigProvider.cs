using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class WindowConfigProvider : ConfigurationServiceBase<WindowSettings>
    {
        public WindowConfigProvider(IPersistenceService<WindowSettings> persister)
            : base(persister, new DefaultSettings<WindowSettings, Properties.Settings>()) { }
    }
}
