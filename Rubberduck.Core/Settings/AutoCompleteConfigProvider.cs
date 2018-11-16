using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class AutoCompleteConfigProvider : IConfigProvider<AutoCompleteSettings>
    {
        private readonly IPersistanceService<AutoCompleteSettings> _persister;
        private readonly AutoCompleteSettings _defaultSettings;

        public AutoCompleteConfigProvider(IPersistanceService<AutoCompleteSettings> persister)
        {
            _persister = persister;
            _defaultSettings = new DefaultSettings<AutoCompleteSettings>().Default;
        }

        public AutoCompleteSettings Create()
        {
            return _persister.Load(_defaultSettings) ?? _defaultSettings;
        }

        public AutoCompleteSettings CreateDefaults()
        {
            return _defaultSettings;
        }

        public void Save(AutoCompleteSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
