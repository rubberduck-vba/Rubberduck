using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using Rubberduck.UnitTesting.Settings;
using UnitTestSettings = Rubberduck.UnitTesting.Settings.UnitTestSettings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class UnitTestSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var unitTestSettings = new UnitTestSettings(BindingMode.LateBinding, AssertMode.StrictAssert, true, true, false);

            var userSettings = new UserSettings(null, null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var unitTestSettings = new UnitTestSettings(BindingMode.EarlyBinding, AssertMode.PermissiveAssert, false, false, true);

            var userSettings = new UserSettings(null, null, null, null, null, unitTestSettings, null, null);
            return new Configuration(userSettings);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(customConfig, null);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(config.UserSettings.UnitTestSettings.BindingMode, viewModel.BindingMode);
                Assert.AreEqual(config.UserSettings.UnitTestSettings.AssertMode, viewModel.AssertMode);
                Assert.AreEqual(config.UserSettings.UnitTestSettings.ModuleInit, viewModel.ModuleInit);
                Assert.AreEqual(config.UserSettings.UnitTestSettings.MethodInit, viewModel.MethodInit);
                Assert.AreEqual(config.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.DefaultTestStubInNewModule);
            });
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new UnitTestSettingsViewModel(GetNondefaultConfig(), null);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.BindingMode, viewModel.BindingMode);
                Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.AssertMode, viewModel.AssertMode);
                Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.ModuleInit, viewModel.ModuleInit);
                Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.MethodInit, viewModel.MethodInit);
                Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.DefaultTestStubInNewModule);
            });
        }

        [Category("Settings")]
        [Test]
        public void BindingModeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.BindingMode, viewModel.BindingMode);
        }

        [Category("Settings")]
        [Test]
        public void AssertModeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.AssertMode, viewModel.AssertMode);
        }

        [Category("Settings")]
        [Test]
        public void ModuleInitIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.ModuleInit, viewModel.ModuleInit);
        }

        [Category("Settings")]
        [Test]
        public void MethodInitIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.MethodInit, viewModel.MethodInit);
        }

        [Category("Settings")]
        [Test]
        public void DefaultTestStubInNewModuleIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new UnitTestSettingsViewModel(defaultConfig, null);

            Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.DefaultTestStubInNewModule);
        }
    }
}
