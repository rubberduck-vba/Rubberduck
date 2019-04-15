using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.UnitTesting.Settings
{
    public class UnitTestConfigProvider : ConfigurationServiceBase<UnitTestSettings>
    {
        public UnitTestConfigProvider(IPersistenceService<UnitTestSettings> persister)
            : base(persister, new DefaultSettings<UnitTestSettings, Properties.UnitTestDefaults>())
        {
        }
    }
}
