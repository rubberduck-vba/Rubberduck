using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class InspectionProviderTests
    {
        [Category("Inspections")]
        [Test]
        public void InspectionTypeIsAssignedFromDefaultSettingInConstructor()
        {
            var defaultSettings = new DefaultSettings<CodeInspectionSettings>().Default;
            var defaultSetting = defaultSettings.CodeInspections.First();
            defaultSetting.InspectionType = CodeInspectionType.Performance;

            var inspectionMock = new Mock<IInspection>();
            inspectionMock.Setup(inspection => inspection.Name).Returns(defaultSetting.Name);
            inspectionMock.Setup(inspection => inspection.InspectionType).Returns(CodeInspectionType.CodeQualityIssues);

            new InspectionProvider(new[] {inspectionMock.Object});

            inspectionMock.VerifySet(inspection => inspection.InspectionType = CodeInspectionType.Performance);
        }
    }
}
