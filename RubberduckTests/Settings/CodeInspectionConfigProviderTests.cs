using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class CodeInspectionConfigProviderTests
    {
        [Category("Settings")]
        [Test]
        public void SettingsForFoundInspectionsAreAddedToDefaultSettings()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock.Setup(inspection => inspection.Name).Returns(inspectionMock.Object.GetType().FullName);
            var configProvider = new CodeInspectionConfigProvider(null, new[] {inspectionMock.Object});

            var defaults = configProvider.CreateDefaults();

            Assert.NotNull(defaults.GetSetting(inspectionMock.Object.GetType()));
        }
    }
}
