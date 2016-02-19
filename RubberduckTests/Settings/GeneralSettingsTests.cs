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
                    new Hotkey{Name="IndentProcedure", IsEnabled=true, KeyDisplaySymbol="CTRL-P"},
                    new Hotkey{Name="IndentModule", IsEnabled=true, KeyDisplaySymbol="CTRL-M"}
                }
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
                    new Hotkey{Name="IndentProcedure", IsEnabled=false, KeyDisplaySymbol="CTRL-C"},
                    new Hotkey{Name="IndentModule", IsEnabled=false, KeyDisplaySymbol="CTRL-X"}
                }
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
                () => Assert.IsTrue(config.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(viewModel.Hotkeys)));
        }

        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new GeneralSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage),
                () => Assert.IsTrue(defaultConfig.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(viewModel.Hotkeys)));
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
    }
}