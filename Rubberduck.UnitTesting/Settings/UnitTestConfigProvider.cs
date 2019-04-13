using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.UnitTesting.Settings
{
    public class UnitTestConfigProvider : ConfigurationServiceBase<UnitTestSettings>
    {
        private readonly UnitTestSettings defaultSettings;

        public UnitTestConfigProvider(IPersistanceService<UnitTestSettings> persister)
            : base(persister)
        {
            defaultSettings = new DefaultSettings<UnitTestSettings, Properties.UnitTestDefaults>().Default;
        }

        public override UnitTestSettings Load()
        {
            return persister.Load(defaultSettings) ?? defaultSettings;
        }

        public override UnitTestSettings LoadDefaults()
        {
            return defaultSettings;
        }
    }
}
