using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using UnitTestSettings = Rubberduck.Settings.UnitTestSettings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class UnitTestSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var unitTestSettings = new UnitTestSettings
            {
                BindingMode = BindingMode.LateBinding,
                AssertMode = AssertMode.StrictAssert,
                ModuleInit = true,
                MethodInit = true,
                DefaultTestStubInNewModule = false
            };

            var userSettings = new UserSettings(null, null, null, unitTestSettings, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var unitTestSettings = new UnitTestSettings
            {
                BindingMode = BindingMode.EarlyBinding,
                AssertMode = AssertMode.PermissiveAssert,
                ModuleInit = false,
                MethodInit = false,
                DefaultTestStubInNewModule = true
            };

            var userSettings = new UserSettings(null, null, null, unitTestSettings, null);
            return new Configuration(userSettings);
        }

        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(GetNondefaultConfig());

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(customConfig.UserSettings.UnitTestSettings.BindingMode, viewModel.BindingMode),
                () => Assert.AreEqual(customConfig.UserSettings.UnitTestSettings.AssertMode, viewModel.AssertMode),
                () => Assert.AreEqual(customConfig.UserSettings.UnitTestSettings.ModuleInit, viewModel.ModuleInit),
                () => Assert.AreEqual(customConfig.UserSettings.UnitTestSettings.MethodInit, viewModel.MethodInit),
                () => Assert.AreEqual(customConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.DefaultTestStubInNewModule));
        }

        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new UnitTestSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.BindingMode, viewModel.BindingMode),
                () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.AssertMode, viewModel.AssertMode),
                () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.ModuleInit, viewModel.ModuleInit),
                () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.MethodInit, viewModel.MethodInit),
                () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.DefaultTestStubInNewModule));
        }

        [TestMethod]
        public void BindingModeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.BindingMode, viewModel.BindingMode);
        }

        [TestMethod]
        public void AssertModeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.AssertMode, viewModel.AssertMode);
        }

        [TestMethod]
        public void ModuleInitIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.ModuleInit, viewModel.ModuleInit);
        }

        [TestMethod]
        public void MethodInitIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.MethodInit, viewModel.MethodInit);
        }

        [TestMethod]
        public void DefaultTestStubInNewModuleIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.DefaultTestStubInNewModule);
        }
    }
}