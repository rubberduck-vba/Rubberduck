using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;
using Rubberduck.Common;
using Moq;
using Rubberduck.UI;

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
                EnableExperimentalFeatures = new List<ExperimentalFeatures>
                {
                    new ExperimentalFeatures
                    {
                        Key = nameof(RubberduckUI.GeneralSettings_EnableSourceControl),
                        IsEnabled = true
                    }
                }
                //Delimiter = '.'
            };

            var hotkeySettings = new HotkeySettings(new[]
            {
                new HotkeySetting {CommandTypeName = "FooCommand", IsEnabled = true, Key1 = "A"},
                new HotkeySetting {CommandTypeName = "BarCommand", IsEnabled = true, Key1 = "B"}
            });

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

            var hotkeySettings = new HotkeySettings
            {
                Settings = new[]
                {
                    new HotkeySetting{CommandTypeName="FooCommand", IsEnabled=false, Key1="C"},
                    new HotkeySetting{CommandTypeName="BarCommand", IsEnabled=false, Key1="D"}
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

            Assert.IsTrue(defaultConfig.UserSettings.GeneralSettings.EnableExperimentalFeatures.SequenceEqual(viewModel.ExperimentalFeatures));
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
