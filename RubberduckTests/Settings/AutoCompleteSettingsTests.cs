using System.Collections.Generic;
using System.Linq;
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
                CompleteBlockOnTab = true,
                CompleteBlockOnEnter = true,
                EnableSmartConcat = true,
                AutoCompletes = new HashSet<AutoCompleteSetting>(new[]
                {
                    new AutoCompleteSetting("AutoCompleteClosingBrace", true),
                    new AutoCompleteSetting("AutoCompleteClosingBracket", true),
                    new AutoCompleteSetting("AutoCompleteClosingParenthese", true),
                    new AutoCompleteSetting("AutoCompleteClosingString", true),
                    new AutoCompleteSetting("AutoCompleteDoBlock", true),
                    new AutoCompleteSetting("AutoCompleteEnumBlock", true),
                    new AutoCompleteSetting("AutoCompleteForBlock", true),
                    new AutoCompleteSetting("AutoCompleteFunctionBlock", true),
                    new AutoCompleteSetting("AutoCompleteIfBlock", true),
                    new AutoCompleteSetting("AutoCompleteOnErrorResumeNextBlock", true),
                    new AutoCompleteSetting("AutoCompletePrecompilerIfBlock", true),
                    new AutoCompleteSetting("AutoCompletePropertyBlock", true),
                    new AutoCompleteSetting("AutoCompleteSelectBlock", true),
                    new AutoCompleteSetting("AutoCompleteSubBlock", true),
                    new AutoCompleteSetting("AutoCompleteTypeBlock", true),
                    new AutoCompleteSetting("AutoCompleteWhileBlock", true),
                    new AutoCompleteSetting("AutoCompleteWithBlock", true)
                })
            };

            var userSettings = new UserSettings(null, null, autoCompleteSettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        private static Configuration GetNonDefaultConfig()
        {
            var autoCompleteSettings = new AutoCompleteSettings
            {
                IsEnabled = true,
                CompleteBlockOnTab = false,
                CompleteBlockOnEnter = false,
                EnableSmartConcat = false,
                AutoCompletes = new HashSet<AutoCompleteSetting>(new[]
                {
                    new AutoCompleteSetting("AutoCompleteClosingBrace", false),
                    new AutoCompleteSetting("AutoCompleteClosingBracket", false),
                    new AutoCompleteSetting("AutoCompleteClosingParenthese", false),
                    new AutoCompleteSetting("AutoCompleteClosingString", false),
                    new AutoCompleteSetting("AutoCompleteDoBlock", false),
                    new AutoCompleteSetting("AutoCompleteEnumBlock", false),
                    new AutoCompleteSetting("AutoCompleteForBlock", false),
                    new AutoCompleteSetting("AutoCompleteFunctionBlock", false),
                    new AutoCompleteSetting("AutoCompleteIfBlock", false),
                    new AutoCompleteSetting("AutoCompleteOnErrorResumeNextBlock", false),
                    new AutoCompleteSetting("AutoCompletePrecompilerIfBlock", false),
                    new AutoCompleteSetting("AutoCompletePropertyBlock", false),
                    new AutoCompleteSetting("AutoCompleteSelectBlock", false),
                    new AutoCompleteSetting("AutoCompleteSubBlock", false),
                    new AutoCompleteSetting("AutoCompleteTypeBlock", false),
                    new AutoCompleteSetting("AutoCompleteWhileBlock", false),
                    new AutoCompleteSetting("AutoCompleteWithBlock", false)
                })
            };

            var userSettings = new UserSettings(null, null, autoCompleteSettings, null, null, null, null, null);
            return new Configuration(userSettings);
        }

        // TODO: Remove this once this feature is stable and it can default to enabled.
        [Category("Settings")]
        [Test]
        public void AutoCompleteDisabledByDefault()
        {
            var defaultSettings = new DefaultSettings<AutoCompleteSettings>().Default;
            Assert.IsFalse(defaultSettings.IsEnabled);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNonDefaultConfig();
            var viewModel = new AutoCompleteSettingsViewModel(customConfig);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.IsEnabled, viewModel.IsEnabled);
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.CompleteBlockOnTab, viewModel.CompleteBlockOnTab);
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.CompleteBlockOnEnter, viewModel.CompleteBlockOnEnter);
                Assert.AreEqual(config.UserSettings.AutoCompleteSettings.EnableSmartConcat, viewModel.EnableSmartConcat);
                Assert.IsTrue(config.UserSettings.AutoCompleteSettings.AutoCompletes.SequenceEqual(viewModel.Settings));
            });
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new AutoCompleteSettingsViewModel(GetNonDefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.IsEnabled, viewModel.IsEnabled);
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.CompleteBlockOnTab, viewModel.CompleteBlockOnTab);
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.CompleteBlockOnEnter, viewModel.CompleteBlockOnEnter);
                Assert.AreEqual(defaultConfig.UserSettings.AutoCompleteSettings.EnableSmartConcat, viewModel.EnableSmartConcat);
                Assert.IsTrue(defaultConfig.UserSettings.AutoCompleteSettings.AutoCompletes.SequenceEqual(viewModel.Settings));
            });
        }
    }
}
