using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class GeneralSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var generalSettings = new GeneralSettings
            {
                Language = new DisplayLanguageSetting("en-US"),
                HotkeySettings = new[]
                {
                    new HotkeySetting{Name="IndentProcedure", IsEnabled=true, Key1="CTRL-P"},
                    new HotkeySetting{Name="IndentModule", IsEnabled=true, Key1="CTRL-M"}
                },
                AutoSaveEnabled = false,
                AutoSavePeriod = 10
            };

            var userSettings = new UserSettings(generalSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var generalSettings = new GeneralSettings
            {
                Language = new DisplayLanguageSetting("sv-SE"),
                HotkeySettings = new[]
                {
                    new HotkeySetting{Name="IndentProcedure", IsEnabled=false, Key1="CTRL-C"},
                    new HotkeySetting{Name="IndentModule", IsEnabled=false, Key1="CTRL-X"}
                },
                AutoSaveEnabled = true,
                AutoSavePeriod = 5
            };

            var userSettings = new UserSettings(generalSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new GeneralSettingsViewModel(customConfig);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(config.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage),
                () => Assert.IsTrue(config.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(viewModel.Hotkeys)),
                () => Assert.AreEqual(config.UserSettings.GeneralSettings.AutoSaveEnabled, viewModel.AutoSaveEnabled),
                () => Assert.AreEqual(config.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod));
        }

        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new GeneralSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage),
                () => Assert.IsTrue(defaultConfig.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(viewModel.Hotkeys)),
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, viewModel.AutoSaveEnabled),
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod));
        }

        [TestMethod]
        public void LanguageIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage);
        }

        [TestMethod]
        public void HotkeysAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(viewModel.Hotkeys));
        }

        [TestMethod]
        public void AutoSaveEnabledIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, viewModel.AutoSaveEnabled);
        }

        [TestMethod]
        public void AutoSavePeriodIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod);
        }
    }
}