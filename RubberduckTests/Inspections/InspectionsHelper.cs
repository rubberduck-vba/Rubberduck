﻿using Moq;
using Rubberduck.Inspections.Rubberduck.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Settings;
using System.Linq;
using Rubberduck.Inspections;

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
                Severity = inspection.Severity, 
                Name = inspection.ToString()
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}