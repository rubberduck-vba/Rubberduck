using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public class UnitTestConfigProvider : IConfigProvider<UnitTestSettings>
    {
        private readonly IPersistanceService<UnitTestSettings> _persister;

        public UnitTestConfigProvider(IPersistanceService<UnitTestSettings> persister)
        {
            _persister = persister;
        }

        public UnitTestSettings Create()
        {
            var prototype = new UnitTestSettings();
            return _persister.Load(prototype) ?? prototype;
        }

        public UnitTestSettings CreateDefaults()
        {
            return new UnitTestSettings();
        }

        public void Save(UnitTestSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
