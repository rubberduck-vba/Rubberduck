using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;
using Rubberduck.Common;
using Moq;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class GeneralSettingsTests
    {
        private Mock<IOperatingSystem> GetOperatingSystemMock()
        {
            return new Mock<IOperatingSystem>();
        }

        private Configuration GetDefaultConfig()
        {
            var generalSettings = new GeneralSettings
            {
                Language = new DisplayLanguageSetting("en-US"),
                AutoSaveEnabled = false,
                AutoSavePeriod = 10,
                //Delimiter = '.'
            };

            var hotkeySettings = new HotkeySettings()
            {
                Settings = new[]
                {
                    new HotkeySetting {Name = "IndentProcedure", IsEnabled = true, Key1 = "CTRL-P"},
                    new HotkeySetting {Name = "IndentModule", IsEnabled = true, Key1 = "CTRL-M"}
                }
            };

            var userSettings = new UserSettings(generalSettings, hotkeySettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var generalSettings = new GeneralSettings
            {
                Language = new DisplayLanguageSetting("fr-CA"),
                AutoSaveEnabled = true,
                AutoSavePeriod = 5,
                //Delimiter = '/'
            };

            var hotkeySettings = new HotkeySettings()
            {
                Settings = new[]
                {
                    new HotkeySetting{Name="IndentProcedure", IsEnabled=false, Key1="CTRL-C"},
                    new HotkeySetting{Name="IndentModule", IsEnabled=false, Key1="CTRL-X"}
                }
            };

            var userSettings = new UserSettings(generalSettings, hotkeySettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new GeneralSettingsViewModel(customConfig, GetOperatingSystemMock().Object);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(config.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage),
                () => Assert.IsTrue(config.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys)),
                () => Assert.AreEqual(config.UserSettings.GeneralSettings.AutoSaveEnabled, viewModel.AutoSaveEnabled),
                () => Assert.AreEqual(config.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new GeneralSettingsViewModel(GetNondefaultConfig(), GetOperatingSystemMock().Object);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage),
                () => Assert.IsTrue(defaultConfig.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys)),
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, viewModel.AutoSaveEnabled),
                () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void LanguageIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void HotkeysAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.IsTrue(defaultConfig.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AutoSaveEnabledIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, viewModel.AutoSaveEnabled);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AutoSavePeriodIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod);
        }

        //[TestCategory("Settings")]
        //[TestMethod]
        //public void DelimiterIsSetInCtor()
        //{
        //    var defaultConfig = GetDefaultConfig();
        //    var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

        //    Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Delimiter, (char)viewModel.Delimiter);
        //}
    }
}
