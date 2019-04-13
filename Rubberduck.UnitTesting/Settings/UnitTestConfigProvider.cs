using Rubberduck.SettingsProvider;

namespace Rubberduck.UnitTesting.Settings
{
    public class UnitTestConfigProvider : ConfigurationServiceBase<UnitTestSettings>
    {
        private readonly UnitTestSettings defaultSettings;

        public UnitTestConfigProvider(IPersistanceService<UnitTestSettings> persister)
            : base(persister)
        {

            defaultSettings = new UnitTestSettings
            {
                BindingMode = BindingMode.LateBinding,
                AssertMode = AssertMode.StrictAssert,
                ModuleInit = true,
                MethodInit = true,
                DefaultTestStubInNewModule = false
            };
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
