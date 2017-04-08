using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class InspectionSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var inspectionSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[]
                {
                    new CodeInspectionSetting("DoNotShowInspection", "Do not show me", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.DoNotShow, CodeInspectionSeverity.DoNotShow),
                    new CodeInspectionSetting("HintInspection", "I'm a hint", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint, CodeInspectionSeverity.Hint),
                    new CodeInspectionSetting("SuggestionInspection", "I'm a suggestion", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion, CodeInspectionSeverity.Suggestion),
                    new CodeInspectionSetting("WarningInspection", "I'm a warning", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning, CodeInspectionSeverity.Warning),
                    new CodeInspectionSetting("ErrorInspection", "FIX ME!", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error, CodeInspectionSeverity.Error),
                    new CodeInspectionSetting("NondefaultSeverityInspection", "I do not have my original severity", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning, CodeInspectionSeverity.DoNotShow)
                })
            };

            var userSettings = new UserSettings(null, null, null, inspectionSettings, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var inspectionSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[]
                {
                    new CodeInspectionSetting("DoNotShowInspection", "Do not show me", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.DoNotShow, CodeInspectionSeverity.Warning),
                    new CodeInspectionSetting("HintInspection", "I'm a hint", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint, CodeInspectionSeverity.Suggestion),
                    new CodeInspectionSetting("SuggestionInspection", "I'm a suggestion", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion, CodeInspectionSeverity.Hint),
                    new CodeInspectionSetting("WarningInspection", "I'm a warning", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Warning, CodeInspectionSeverity.Error),
                    new CodeInspectionSetting("ErrorInspection", "FIX ME!", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error, CodeInspectionSeverity.DoNotShow),
                    new CodeInspectionSetting("NondefaultSeverityInspection", "I do not have my original severity", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Warning, CodeInspectionSeverity.Error)
                })
            };

            var userSettings = new UserSettings(null, null, null, inspectionSettings, null, null, null);
            return new Configuration(userSettings);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new InspectionSettingsViewModel(customConfig);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.IsTrue(config.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>()));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new InspectionSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>()));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void InspectionsAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new InspectionSettingsViewModel(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>()));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void InspectionSeveritiesAreUpdated()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new InspectionSettingsViewModel(defaultConfig);

            viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().First().Severity =
                GetNondefaultConfig().UserSettings.CodeInspectionSettings.CodeInspections.First().Severity;

            var updatedConfig = defaultConfig;
            updatedConfig.UserSettings.CodeInspectionSettings.CodeInspections.First().Severity =
                GetNondefaultConfig().UserSettings.CodeInspectionSettings.CodeInspections.First().Severity;

            Assert.IsTrue(updatedConfig.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>()));
        }
    }
}
