using System;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Interaction;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.VersionCheck;

namespace RubberduckTests
{
    [TestFixture]
    public class VersionCheckTests
    {
        private Configuration CreateConfig(bool includePreReleases)
        {
            var general = new GeneralSettings
            {
                CanCheckVersion = true,
                IncludePreRelease = includePreReleases
            };
            var settings = new UserSettings(general, null, null, null, null, null, null, null);
            return new Configuration(settings);
        }

        private VersionCheckCommand CreateSUT(Configuration config, Version currentVersion, Version latestVersion, out Mock<IMessageBox> mockPrompt, out Mock<IVersionCheckService> mockService)
        {
            mockPrompt = new Mock<IMessageBox>();
            mockPrompt.Setup(m => m.Question(It.IsAny<string>(), It.IsAny<string>()));

            var mockProcess = new Mock<IExternalProcess>();
            var mockConfig = new Mock<IConfigurationService<Configuration>>();
            mockConfig.Setup(m => m.Read()).Returns(() => config);

            mockService = new Mock<IVersionCheckService>();
            
            mockService.Setup(m => m.CurrentVersion)
                       .Returns(() => currentVersion);

            mockService.Setup(m => m.GetLatestVersionAsync(It.IsAny<GeneralSettings>(), It.IsAny<CancellationToken>()))
                       .ReturnsAsync(() => latestVersion);

            return new VersionCheckCommand(mockService.Object, mockPrompt.Object, mockProcess.Object, mockConfig.Object);
        }

        [Test]
        public void DebugBuild_ReleaseOnly_LatestHasRevisionNumber_NoPrompt()
        {
            var config = CreateConfig(includePreReleases: false);
            var currentVersion = new Version(2, 5, 2, 0);
            var latestVersion = new Version(2, 5, 2, 1);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => true);

            sut.Execute(null);

            mockPrompt.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Never());
        }
        
        [Test]
        public void DebugBuild_PreReleases_NoPrompt()
        {
            var config = CreateConfig(includePreReleases: true);
            var currentVersion = new Version(2, 5, 2, 0);
            var latestVersion = new Version(2, 5, 2, 5678);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => true);

            sut.Execute(null);

            mockPrompt.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Never());
        }
        
        [Test]
        public void DebugBuild_ReleaseOnly_Prompt()
        {
            var config = CreateConfig(includePreReleases: false);
            var currentVersion = new Version(2, 5, 2, 0);
            var latestVersion = new Version(2, 5, 3, 0);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => true);

            sut.Execute(null);

            mockPrompt.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Never());
        }
        
        [Test]
        public void DebugBuild_PreReleases_Prompt()
        {
            var config = CreateConfig(includePreReleases: true);
            var currentVersion = new Version(2, 5, 2, 0);
            var latestVersion = new Version(2, 5, 3, 5678);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => true);

            sut.Execute(null);

            mockPrompt.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Never());
        }
        
        [Test]
        public void ReleaseBuild_ReleaseOnly_LatestHasRevisionNumber_NoPrompt()
        {
            var config = CreateConfig(includePreReleases: false);
            var currentVersion = new Version(2, 5, 2, 4567);
            var latestVersion = new Version(2, 5, 2, 1);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => false);

            sut.Execute(null);

            mockPrompt.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Never());
        }
        
        [Test]
        public void ReleaseBuild_PreReleases_Prompt()
        {
            var config = CreateConfig(includePreReleases: true);
            var currentVersion = new Version(2, 5, 2, 4567);
            var latestVersion = new Version(2, 5, 2, 5678);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => false);

            sut.Execute(null);

            mockPrompt.Verify();
        }
        
        [Test]
        public void ReleaseBuild_ReleaseOnly_Prompt()
        {
            var config = CreateConfig(includePreReleases: false);
            var currentVersion = new Version(2, 5, 2, 4567);
            var latestVersion = new Version(2, 5, 3, 0);

            var sut = CreateSUT(config, currentVersion, latestVersion, out var mockPrompt, out var mockService);
            mockService.Setup(m => m.IsDebugBuild).Returns(() => false);

            sut.Execute(null);

            mockPrompt.Verify();
        }
    }
}