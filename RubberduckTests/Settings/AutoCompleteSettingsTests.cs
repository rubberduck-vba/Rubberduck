using NUnit.Framework;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using AutoCompleteSettings = Rubberduck.Settings.AutoCompleteSettings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class AutoCompleteSettingsTests
    {
        private static Configuration GetDefaultConfig()
        {
            var autoCompleteSettings = new AutoCompleteSettings
            {
                IsEnabled = false,
                BlockCompletion = new AutoCompleteSettings.BlockCompletionSettings
                {
                    CompleteOnTab = true,
                    CompleteOnEnter = true,
                    IsEnabled = true
                },
                SmartConcat = new AutoCompleteSettings.SmartConcatSettings
                {
                    IsEnabled = true,
                    ConcatVbNewLineModifier = ModifierKeySetting.CtrlKey
                },
                SelfClosingPairs = new AutoCompleteSettings.SelfClosingPairSettings
                {
                    IsEnabled = true
                }
            };

            var userSettings = new UserSettings(null, null, autoCompleteSettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        private static Configuration GetNonDefaultConfig()
        {
            var autoCompleteSettings = new AutoCompleteSettings
            {
                IsEnabled = true,
                BlockCompletion = new AutoCompleteSettings.BlockCompletionSettings
                {
                    CompleteOnTab = false,
                    CompleteOnEnter = false,
                    IsEnabled = false
                },
                SmartConcat = new AutoCompleteSettings.SmartConcatSettings
                {
                    IsEnabled = false,
                    ConcatVbNewLineModifier = ModifierKeySetting.CtrlKey
                },
                SelfClosingPairs = new AutoCompleteSettings.SelfClosingPairSettings
                {
                    IsEnabled = false
                }

            };

            var userSettings = new UserSettings(null, null, autoCompleteSettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        // TODO: Remove this once this feature is stable and it can default to enabled.
        [Category("Settings")]
        [Test]
        public void AutoCompleteDisabledByDefault()
        {
            var defaultSettings = new DefaultSettings<AutoCompleteSettings, Rubberduck.Properties.Settings>().Default;
            Assert.IsFalse(defaultSettings.IsEnabled);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNonDefaultConfig();
            var viewModel = new AutoCompleteSettingsViewModel(customConfig, null);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.IsEnabled, viewModel.IsEnabled);
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.BlockCompletion.CompleteOnTab, viewModel.CompleteBlockOnTab);
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.BlockCompletion.CompleteOnEnter, viewModel.CompleteBlockOnEnter);
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.SmartConcat.IsEnabled, viewModel.EnableSmartConcat);
            });
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new AutoCompleteSettingsViewModel(GetNonDefaultConfig(), null);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.IsEnabled, viewModel.IsEnabled);
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.BlockCompletion.CompleteOnTab, viewModel.CompleteBlockOnTab);
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.BlockCompletion.CompleteOnEnter, viewModel.CompleteBlockOnEnter);
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.SmartConcat.IsEnabled, viewModel.EnableSmartConcat);
            });
        }
    }
}
