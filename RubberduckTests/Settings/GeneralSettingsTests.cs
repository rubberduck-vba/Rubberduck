using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;
using Rubberduck.Common;
using Moq;

namespace RubberduckTests.Settings
{
    [TestFixture]
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
                IsAutoSaveEnabled = false,
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
                IsAutoSaveEnabled = true,
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

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new GeneralSettingsViewModel(customConfig, GetOperatingSystemMock().Object);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(config.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage);
                Assert.IsTrue(config.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys));
                Assert.AreEqual(config.UserSettings.GeneralSettings.IsAutoSaveEnabled, viewModel.AutoSaveEnabled);
                Assert.AreEqual(config.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod);
            });
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new GeneralSettingsViewModel(GetNondefaultConfig(), GetOperatingSystemMock().Object);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage);
                Assert.IsTrue(defaultConfig.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys));
                Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.IsAutoSaveEnabled, viewModel.AutoSaveEnabled);
                Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod);
            });
        }

        [Category("Settings")]
        [Test]
        public void LanguageIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage);
        }

        [Category("Settings")]
        [Test]
        public void HotkeysAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.IsTrue(defaultConfig.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys));
        }

        [Category("Settings")]
        [Test]
        public void AutoSaveEnabledIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.IsAutoSaveEnabled, viewModel.AutoSaveEnabled);
        }

        [Category("Settings")]
        [Test]
        public void AutoSavePeriodIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod);
        }

        [Category("Settings")]
        [Test]
        public void SourceControlEnabledIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.IsSourceControlEnabled, viewModel.SourceControlEnabled);
        }

        //[Category("Settings")]
        //[Test]
        //public void DelimiterIsSetInCtor()
        //{
        //    var defaultConfig = GetDefaultConfig();
        //    var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object);

        //    Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Delimiter, (char)viewModel.Delimiter);
        //}
    }
}
