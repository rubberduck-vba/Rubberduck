using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Logistics;
using Rubberduck.SettingsProvider;
using Rubberduck.CodeAnalysis.Settings;

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

            var defaults = configProvider.ReadDefaults();

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

            var persisterMock = new Mock<IPersistenceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(null)).Returns(userSettings);

            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, inspectionProviderMock.Object);

            var settings = configProvider.Read().CodeInspections;
            var defaultSettings = configProvider.ReadDefaults().CodeInspections;

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

            var persisterMock = new Mock<IPersistenceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(null)).Returns(userSettings);

            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, inspectionProviderMock.Object);
            var settings = configProvider.Read().CodeInspections;

            Assert.AreEqual(configProvider.ReadDefaults().CodeInspections.Count, settings.Count);
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

            var persisterMock = new Mock<IPersistenceService<CodeInspectionSettings>>();
            persisterMock.Setup(persister => persister.Load(null)).Returns(userSettings);

            var configProvider = new CodeInspectionConfigProvider(persisterMock.Object, inspectionProviderMock.Object);

            var settings = configProvider.Read().CodeInspections;

            Assert.IsNull(settings.FirstOrDefault(setting => setting.Name == "Bar"));
        }
    }
}
