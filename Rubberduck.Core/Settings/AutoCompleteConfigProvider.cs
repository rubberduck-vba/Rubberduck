using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class AutoCompleteConfigProvider : ConfigurationServiceBase<AutoCompleteSettings>
    {
        public AutoCompleteConfigProvider(IPersistenceService<AutoCompleteSettings> persister)
            : base(persister, new DefaultSettings<AutoCompleteSettings, Properties.Settings>()) { }
    }
}
