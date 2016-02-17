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
            var unitTestSettings = new UnitTestSettings()
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
            var unitTestSettings = new UnitTestSettings()
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
            var viewModel = new UnitTestSettingsViewModel(GetNondefaultConfig());
            viewModel.UpdateConfig(GetNondefaultConfig());

            Assert.AreEqual(BindingMode.EarlyBinding, viewModel.BindingMode);
            Assert.AreEqual(AssertMode.PermissiveAssert, viewModel.AssertMode);
            Assert.AreEqual(false, viewModel.ModuleInit);
            Assert.AreEqual(false, viewModel.MethodInit);
            Assert.AreEqual(true, viewModel.DefaultTestStubInNewModule);
        }

        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new UnitTestSettingsViewModel(GetNondefaultConfig());

            viewModel.SetToDefaults(GetDefaultConfig());

            Assert.AreEqual(BindingMode.LateBinding, viewModel.BindingMode);
            Assert.AreEqual(AssertMode.StrictAssert, viewModel.AssertMode);
            Assert.AreEqual(true, viewModel.ModuleInit);
            Assert.AreEqual(true, viewModel.MethodInit);
            Assert.AreEqual(false, viewModel.DefaultTestStubInNewModule);
        }

        [TestMethod]
        public void BindingModeIsSetInCtor()
        {
            var viewModel = new UnitTestSettingsViewModel(GetDefaultConfig());

            Assert.AreEqual(BindingMode.LateBinding, viewModel.BindingMode);
        }

        [TestMethod]
        public void AssertModeIsSetInCtor()
        {
            var viewModel = new UnitTestSettingsViewModel(GetDefaultConfig());

            Assert.AreEqual(AssertMode.StrictAssert, viewModel.AssertMode);
        }

        [TestMethod]
        public void ModuleInitIsSetInCtor()
        {
            var viewModel = new UnitTestSettingsViewModel(GetDefaultConfig());

            Assert.AreEqual(true, viewModel.ModuleInit);
        }

        [TestMethod]
        public void MethodInitIsSetInCtor()
        {
            var viewModel = new UnitTestSettingsViewModel(GetDefaultConfig());

            Assert.AreEqual(true, viewModel.MethodInit);
        }

        [TestMethod]
        public void DefaultTestStubInNewModuleIsSetInCtor()
        {
            var viewModel = new UnitTestSettingsViewModel(GetDefaultConfig());

            Assert.AreEqual(false, viewModel.DefaultTestStubInNewModule);
        }
    }
}