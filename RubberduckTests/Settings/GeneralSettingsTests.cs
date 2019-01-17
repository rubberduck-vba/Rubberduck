using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.Settings.GeneralSettings;
using Rubberduck.Common;
using Moq;
using Rubberduck.VBEditor.VbeRuntime.Settings;
using System;
using Rubberduck.Interaction;
using Rubberduck.SettingsProvider;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class GeneralSettingsTests
    {
        private Mock<IOperatingSystem> GetOperatingSystemMock()
        {
            return new Mock<IOperatingSystem>();
        }

        private Mock<IMessageBox> GetMessageBoxMock()
        {
            return new Mock<IMessageBox>();
        }

        private Mock<IVbeSettings> GetVbeSettingsMock()
        {
            return new Mock<IVbeSettings>();
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
                    new ExperimentalFeatures()
                }
                //Delimiter = '.'
            };

            var hotkeySettings = new HotkeySettings(new[]
            {
                new HotkeySetting {CommandTypeName = "FooCommand", IsEnabled = true, Key1 = "A"},
                new HotkeySetting {CommandTypeName = "BarCommand", IsEnabled = true, Key1 = "B"}
            });

            var userSettings = new UserSettings(generalSettings, hotkeySettings, null, null, null, null, null, null);
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

            var userSettings = new UserSettings(generalSettings, hotkeySettings, null, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new GeneralSettingsViewModel(customConfig, GetOperatingSystemMock().Object, GetMessageBoxMock().Object, GetVbeSettingsMock().Object, new List<Type>());

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
            var viewModel = new GeneralSettingsViewModel(GetNondefaultConfig(), GetOperatingSystemMock().Object, GetMessageBoxMock().Object, GetVbeSettingsMock().Object, new List<Type>());

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
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object, GetMessageBoxMock().Object, GetVbeSettingsMock().Object, new List<Type>());

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, viewModel.SelectedLanguage);
        }

        [Category("Settings")]
        [Test]
        public void HotkeysAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object, GetMessageBoxMock().Object, GetVbeSettingsMock().Object, new List<Type>());

            Assert.IsTrue(defaultConfig.UserSettings.HotkeySettings.Settings.SequenceEqual(viewModel.Hotkeys));
        }

        [Category("Settings")]
        [Test]
        public void AutoSaveEnabledIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object, GetMessageBoxMock().Object, GetVbeSettingsMock().Object, new List<Type>());

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.IsAutoSaveEnabled, viewModel.AutoSaveEnabled);
        }

        [Category("Settings")]
        [Test]
        public void AutoSavePeriodIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new GeneralSettingsViewModel(defaultConfig, GetOperatingSystemMock().Object, GetMessageBoxMock().Object, GetVbeSettingsMock().Object, new List<Type>());

            Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, viewModel.AutoSavePeriod);
        }

        [Category("Settings")]
        [Test]
        public void UserSettingsLoadedUsingDefaultWhenMissingFile()
        {
            // For this test, we need to use the actual object. Fortunately, the path is virtual, so we
            // can override that property and force it to use an non-existent path to prove that settings
            // will be still created using defaults without the file present. 
            var persisterMock = new Mock<XmlPersistanceService<GeneralSettings>>();
            persisterMock.Setup(x => x.FilePath).Returns("C:\\some\\non\\existent\\path\\rubberduck");
            persisterMock.CallBase = true;
            var configProvider = new GeneralConfigProvider(persisterMock.Object);

            var settings = configProvider.Create();
            var defaultSettings = configProvider.CreateDefaults();

            Assert.AreEqual(defaultSettings, settings);
        }
    }
}
