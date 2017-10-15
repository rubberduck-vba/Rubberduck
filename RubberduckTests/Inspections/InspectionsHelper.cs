using Moq;
using Rubberduck.Inspections.Rubberduck.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Settings;
using System.Linq;

namespace RubberduckTests.Inspections
{
    public static class InspectionsHelper
    {
        public static IInspector GetInspector(IInspection inspection, params IInspection[] otherInspections)
        {
            return new Inspector(GetSettings(inspection), otherInspections.Union(new[] {inspection}));
        }

        public static IGeneralConfigService GetSettings(IInspection inspection)
        {
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig(inspection);
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            return settings.Object;
        }

        private static Configuration GetTestConfig(IInspection inspection)
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = inspection.Description,
                Severity = inspection.Severity
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}