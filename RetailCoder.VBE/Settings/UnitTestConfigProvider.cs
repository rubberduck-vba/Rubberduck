using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public interface IUnitTestConfigProvider
    {
        UnitTestSettings Create();
        UnitTestSettings CreateDefaults();

        void Save(UnitTestSettings settings);
    }

    public class UnitTestConfigProvider : IUnitTestConfigProvider
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
