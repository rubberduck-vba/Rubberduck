using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Inspections;
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
            var inspectionProviderMock = new Mock<IInspectionProvider>();
            inspectionProviderMock.Setup(provider => provider.Inspections).Returns(new[] {inspectionMock.Object});

            var configProvider = new CodeInspectionConfigProvider(null, inspectionProviderMock.Object);

            var defaults = configProvider.CreateDefaults();

            Assert.NotNull(defaults.GetSetting(inspectionMock.Object.GetType()));
        }

        [Category("Settings")]
        [Test]
        public void UserSettingsAreCombinedWithDefaultSettings()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock.Setup(inspection => inspection.Name).Returns("Foo");
            var inspectionProviderMock = new Mock<IInspectionProvider>();
            inspectionProviderMock.Setup(provider => provider.Inspections).Returns(new[] { inspectionMock.Object });

            var userSetting = new CodeInspectionSetting("Foo", CodeInspectionType.CodeQualityIssues);
            var userSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[] { userSetting })
            };

            var persisterMock = new Mock<IPersistanceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(It.IsAny<CodeInspectionSettings>())).Returns(userSettings);

            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, inspectionProviderMock.Object);

            var settings = configProvider.Create().CodeInspections;
            var defaultSettings = configProvider.CreateDefaults().CodeInspections;

            Assert.Contains(userSetting, settings.ToArray());
            Assert.IsTrue(defaultSettings.All(s => settings.Contains(s)));
        }

        [Category("Settings")]
        [Test]
        public void UserSettingsAreNotDuplicatedWithDefaultSettings()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock.Setup(inspection => inspection.Name).Returns("Foo");
            var inspectionProviderMock = new Mock<IInspectionProvider>();
            inspectionProviderMock.Setup(provider => provider.Inspections).Returns(new[] { inspectionMock.Object });

            var userSetting = new CodeInspectionSetting(inspectionMock.Object.Name, inspectionMock.Object.InspectionType);
            var userSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[] { userSetting })
            };

            var persisterMock = new Mock<IPersistanceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(It.IsAny<CodeInspectionSettings>())).Returns(userSettings);

            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, inspectionProviderMock.Object);
            var settings = configProvider.Create().CodeInspections;

            Assert.AreEqual(configProvider.CreateDefaults().CodeInspections.Count, settings.Count);
        }


        [Category("Settings")]
        [Test]
        public void UserSettingForUnknownInspectionIsIgnored()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock.Setup(inspection => inspection.Name).Returns("Foo");
            var inspectionProviderMock = new Mock<IInspectionProvider>();
            inspectionProviderMock.Setup(provider => provider.Inspections).Returns(new[] {inspectionMock.Object});

            var userSetting = new CodeInspectionSetting("Bar", CodeInspectionType.CodeQualityIssues);
            var userSettings = new CodeInspectionSettings
            {
                CodeInspections = new HashSet<CodeInspectionSetting>(new[] { userSetting })
            };

            var persisterMock = new Mock<IPersistanceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(It.IsAny<CodeInspectionSettings>())).Returns(userSettings);

            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, inspectionProviderMock.Object);

            var settings = configProvider.Create().CodeInspections;

            Assert.IsNull(settings.FirstOrDefault(setting => setting.Name == "Bar"));
        }
    }
}
