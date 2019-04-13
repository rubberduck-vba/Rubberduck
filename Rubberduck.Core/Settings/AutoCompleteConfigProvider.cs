using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class AutoCompleteConfigProvider : ConfigurationServiceBase<AutoCompleteSettings>
    {
        private readonly AutoCompleteSettings _defaultSettings;

        public AutoCompleteConfigProvider(IPersistanceService<AutoCompleteSettings> persister)
            : base(persister)
        {
            _defaultSettings = new DefaultSettings<AutoCompleteSettings, Properties.Settings>().Default;
        }

        public override AutoCompleteSettings Load()
        {
            return persister.Load(_defaultSettings) ?? _defaultSettings;
        }

        public override AutoCompleteSettings LoadDefaults()
        {
            return _defaultSettings;
        }
    }
}
