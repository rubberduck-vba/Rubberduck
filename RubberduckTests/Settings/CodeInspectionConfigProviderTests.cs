using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

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

        [Category("Settings")]
        [Test]
        public void UserSettingsAreCombinedWithDefaultSettings()
        {
            var userSetting = new CodeInspectionSetting("Foo", CodeInspectionType.CodeQualityIssues);
            var userSettings = new CodeInspectionSettings { CodeInspections = new HashSet<CodeInspectionSetting>(new[] {userSetting})};
            var persisterMock = new Mock<IPersistanceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(It.IsAny<CodeInspectionSettings>())).Returns(userSettings);
            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, Enumerable.Empty<IInspection>());

            var settings = configProvider.Create().CodeInspections;
            var defaultSettings = configProvider.CreateDefaults().CodeInspections;

            Assert.Contains(userSetting, settings.ToArray());
            Assert.IsTrue(defaultSettings.All(s => settings.Contains(s)));
        }

        [Category("Settings")]
        [Test]
        public void UserSettingsAreNotDuplicatedWithDefaultSettings()
        {
            var defaultSettings = new CodeInspectionConfigProvider(null, Enumerable.Empty<IInspection>()).CreateDefaults().CodeInspections;
            var userSetting = defaultSettings.First();
            var userSettings = new CodeInspectionSettings { CodeInspections = new HashSet<CodeInspectionSetting>(new[] { userSetting }) };
            var persisterMock = new Mock<IPersistanceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(It.IsAny<CodeInspectionSettings>())).Returns(userSettings);
            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, Enumerable.Empty<IInspection>());

            var settings = configProvider.Create().CodeInspections;

            Assert.AreEqual(defaultSettings.Count, settings.Count);
            Assert.Contains(userSetting, settings.ToArray());
        }
    }
}
