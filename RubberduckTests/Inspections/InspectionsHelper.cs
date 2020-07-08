using Moq;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Logistics;
using Rubberduck.SettingsProvider;
using Rubberduck.CodeAnalysis.Settings;

namespace RubberduckTests.Inspections
{
    public static class InspectionsHelper
    {
        public static IInspector GetInspector(IInspection inspection, params IInspection[] otherInspections)
        {
            var inspectionProviderMock = new Mock<IInspectionProvider>();
            inspectionProviderMock.Setup(provider => provider.Inspections).Returns(otherInspections.Union(new[] {inspection}));

            return new Inspector(GetSettings(inspection), inspectionProviderMock.Object);
        }

        public static IConfigurationService<CodeInspectionSettings> GetSettings(IInspection inspection)
        {
            var settings = new Mock<IConfigurationService<CodeInspectionSettings>>();
            var config = GetTestConfig(inspection);
            settings.Setup(x => x.Read()).Returns(config);

            return settings.Object;
        }

        private static CodeInspectionSettings GetTestConfig(IInspection inspection)
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = inspection.Description,
                Severity = inspection.Severity, 
                Name = inspection.ToString()
            });
            return settings;
        }
    }
}