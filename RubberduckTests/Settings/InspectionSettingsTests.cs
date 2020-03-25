using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class InspectionSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var inspectionSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[]
                {
                    new CodeInspectionSetting("DoNotShowInspection", "Do not show me", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.DoNotShow),
                    new CodeInspectionSetting("HintInspection", "I'm a hint", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Hint),
                    new CodeInspectionSetting("SuggestionInspection", "I'm a suggestion", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Suggestion),
                    new CodeInspectionSetting("WarningInspection", "I'm a warning", CodeInspectionType.CodeQualityIssues),
                    new CodeInspectionSetting("NondefaultSeverityInspection", "I do not have my original severity", CodeInspectionType.LanguageOpportunities,CodeInspectionSeverity.DoNotShow),
                    new CodeInspectionSetting("ErrorInspection", "FIX ME!", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error)
                }.OrderBy(cis => cis.TypeLabel)
                    .ThenBy(cis => cis.Description)) // Explicit sorting is to match InspectionSettingsViewModel.cs
            };

            var userSettings = new UserSettings(null, null, null, null, inspectionSettings, null, null, null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var inspectionSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[]
                {
                    new CodeInspectionSetting("DoNotShowInspection", "Do not show me", CodeInspectionType.LanguageOpportunities),
                    new CodeInspectionSetting("HintInspection", "I'm a hint", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Suggestion),
                    new CodeInspectionSetting("SuggestionInspection", "I'm a suggestion", CodeInspectionType.MaintainabilityAndReadabilityIssues, CodeInspectionSeverity.Hint),
                    new CodeInspectionSetting("WarningInspection", "I'm a warning", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.Error),
                    new CodeInspectionSetting("NondefaultSeverityInspection", "I do not have my original severity", CodeInspectionType.LanguageOpportunities, CodeInspectionSeverity.Error),
                    new CodeInspectionSetting("ErrorInspection", "FIX ME!", CodeInspectionType.CodeQualityIssues, CodeInspectionSeverity.DoNotShow)
                }.OrderBy(cis => cis.TypeLabel)
                    .ThenBy(cis => cis.Description)) // Explicit sorting is to match InspectionSettingsViewModel.cs
            };

            var userSettings = new UserSettings(null, null, null, null, inspectionSettings, null, null, null);
            return new Configuration(userSettings);
        }

        [Category("Settings")]
        [Test]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new InspectionSettingsViewModel(customConfig, null);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            Assert.IsTrue(config.UserSettings.CodeInspectionSettings.CodeInspections.OrderBy(setting => setting.InspectionType).SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().OrderBy(setting => setting.InspectionType)));
        }

        [Category("Settings")]
        [Test]
        public void SetDefaultsWorks()
        {
            var viewModel = new InspectionSettingsViewModel(GetNondefaultConfig(), null);

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            Assert.IsTrue(defaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.OrderBy(setting => setting.InspectionType).SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().OrderBy(setting => setting.InspectionType)));
        }

        [Category("Settings")]
        [Test]
        public void InspectionsAreSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new InspectionSettingsViewModel(defaultConfig, null);

            Assert.IsTrue(defaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.OrderBy(setting => setting.InspectionType).SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().OrderBy(setting => setting.InspectionType)));
        }

        [Category("Settings")]
        [Test]
        public void InspectionSeveritiesAreUpdated()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new InspectionSettingsViewModel(defaultConfig, null);

            viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().First().Severity =
                GetNondefaultConfig().UserSettings.CodeInspectionSettings.CodeInspections.First().Severity;

            var updatedConfig = defaultConfig;
            updatedConfig.UserSettings.CodeInspectionSettings.CodeInspections.First().Severity =
                GetNondefaultConfig().UserSettings.CodeInspectionSettings.CodeInspections.First().Severity;

            Assert.IsTrue(updatedConfig.UserSettings.CodeInspectionSettings.CodeInspections.OrderBy(setting => setting.InspectionType).SequenceEqual(
                    viewModel.InspectionSettings.SourceCollection.OfType<CodeInspectionSetting>().OrderBy(setting => setting.InspectionType)));
        }
    }
}
