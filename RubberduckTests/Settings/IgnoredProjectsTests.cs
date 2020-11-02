using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Settings;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestFixture]
    public class IgnoredProjectsTests
    {
        [Test]
        [Category("Settings")]
        public void ConfigProvider_DefaultHasEmptyListOfIgnoredProjects()
        {
            var configProvider = new IgnoredProjectsConfigProvider(new Mock<IPersistenceService<IgnoredProjectsSettings>>().Object);

            var defaultSettings = configProvider.ReadDefaults();

            Assert.False(defaultSettings.IgnoredProjectPaths.Any());
        }


        [Test]
        [Category("Settings")]
        public void ViewModel_InitiallyHasIgnoredPathsFromCurrentSettingsFromProvider()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings);

            var expectedPaths = initialFolders;

            Assert.IsTrue(expectedPaths.SequenceEqual(viewModel.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void RemoveFilename_RemovesExistingFilename()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings);
            viewModel.RemoveSelectedProjects.Execute("eftgsrghdg");

            var expectedPaths = initialFolders.Where(name => !"eftgsrghdg".Equals(name)).ToList();

            Assert.IsTrue(expectedPaths.SequenceEqual(viewModel.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void RemoveFilename_DoesNothingForNonExistingFilename()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings);
            viewModel.RemoveSelectedProjects.Execute("ssss");

            var expectedPaths = initialFolders;

            Assert.IsTrue(expectedPaths.SequenceEqual(viewModel.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void RemoveFilename_RemovesExistingFilenames()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings);
            viewModel.RemoveSelectedProjects.Execute(new[]{"eftgsrghdg", "asdasdasd", "dddd"});

            var expectedPaths = new List<string>{ "faffvseafuoeagfwvef" };

            Assert.IsTrue(expectedPaths.SequenceEqual(viewModel.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void RemoveFilename_DoesNotChangeSettings()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings);
            viewModel.RemoveSelectedProjects.Execute("eftgsrghdg");

            var expectedPaths = initialFolders;

            Assert.IsTrue(expectedPaths.SequenceEqual(initialSettings.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void AddIgnoredFile_AddsFileReturnedFromFileBrowser()
        {
            var initialSettings = new IgnoredProjectsSettings();
            var viewModel = GetSettingsViewModel(initialSettings, "someFilename");
            viewModel.AddIgnoredFileCommand.Execute(null);

            var expectedFolders = new List<string> { "someFilename" };

            Assert.IsTrue(expectedFolders.SequenceEqual(viewModel.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void AddIgnoredFile_DoesNotAddIfNotOK()
        {
            var initialSettings = new IgnoredProjectsSettings();
            var mockFileDialog = FileDialogMock("someFilename", dialogOk: false);
            var mockFileSystemBrowser = FileSystemBrowserMock(mockFileDialog.Object, "asdawedefde");
            var viewModel = GetSettingsViewModel(mockFileSystemBrowser.Object, initialSettings);
            viewModel.AddIgnoredFileCommand.Execute(null);

            Assert.IsFalse(viewModel.IgnoredProjectPaths.Any());
        }

        [Test]
        [Category("Settings")]
        public void AddIgnoredFile_DoesNotAddIfAlreadyThere()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings, "asdasdasd");
            viewModel.AddIgnoredFileCommand.Execute(null);

            var expectedPaths = initialFolders;

            Assert.IsTrue(expectedPaths.SequenceEqual(viewModel.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void AddIgnoredFile_DoesNotChangeSettings()
        {
            var initialSettings = new IgnoredProjectsSettings();
            var viewModel = GetSettingsViewModel(initialSettings, "someFilename");
            viewModel.AddIgnoredFileCommand.Execute(null);

            Assert.IsFalse(initialSettings.IgnoredProjectPaths.Any());
        }

        [Test]
        [Category("Settings")]
        public void UpdateConfig_CallsSave()
        {
            var provider = ConfigProviderMock();
            provider.Setup(m => m.Save(It.IsAny<IgnoredProjectsSettings>()));
            var viewModel = GetSettingsViewModel(provider.Object);

            viewModel.UpdateConfig(null);
            provider.Verify(m => m.Save(It.IsAny<IgnoredProjectsSettings>()), Times.Once);
        }
        
        [Test]
        [Category("Settings")]
        public void UpdateConfig_UsesLoadedSettingsInstance()
        {
            var initialSettings = new IgnoredProjectsSettings();
            var provider = ConfigProviderMock(initialSettings);
            var viewModel = GetSettingsViewModel(provider.Object);

            viewModel.UpdateConfig(null);
            provider.Verify(m => m.Save(initialSettings), Times.Once);
        }

        [Test]
        [Category("Settings")]
        public void UpdateConfig_TransfersIgnoredPathsToSettings()
        {
            var initialSettings = new IgnoredProjectsSettings();
            var viewModel = GetSettingsViewModel(initialSettings, "someFilename");
            viewModel.AddIgnoredFileCommand.Execute(null);

            viewModel.UpdateConfig(null);

            var expectedFolders = new List<string> { "someFilename" };

            Assert.IsTrue(expectedFolders.SequenceEqual(initialSettings.IgnoredProjectPaths));
        }

        [Test]
        [Category("Settings")]
        public void SetToDefaults_ClearsIgnoredPaths()
        {
            var initialFolders = new List<string>
            {
                "asdasdasd",
                "eftgsrghdg",
                "faffvseafuoeagfwvef"
            };
            var initialSettings = new IgnoredProjectsSettings
            {
                IgnoredProjectPaths = initialFolders
            };
            var viewModel = GetSettingsViewModel(initialSettings);

            viewModel.SetToDefaults(null);

            Assert.IsFalse(viewModel.IgnoredProjectPaths.Any());
        }


        private static IgnoredProjectsSettingsViewModel GetSettingsViewModel(IgnoredProjectsSettings initialSettings = null, string newIgnoredFilename = null)
        {
            var providerMock = ConfigProviderMock(initialSettings);
            return GetSettingsViewModel(providerMock.Object, newIgnoredFilename);
        }

        private static IgnoredProjectsSettingsViewModel GetSettingsViewModel(IConfigurationService<IgnoredProjectsSettings> provider, string newIgnoredFilename = null)
        {
            var fileSystemMock = FileSystemBrowserMock(newIgnoredFilename);
            return GetSettingsViewModel(provider, fileSystemMock.Object);
        }

        private static IgnoredProjectsSettingsViewModel GetSettingsViewModel(IFileSystemBrowserFactory fileSystem, IgnoredProjectsSettings initialSettings = null)
        {
            var providerMock = ConfigProviderMock(initialSettings);
            return GetSettingsViewModel(providerMock.Object, fileSystem);
        }

        private static IgnoredProjectsSettingsViewModel GetSettingsViewModel(IConfigurationService<IgnoredProjectsSettings> provider, IFileSystemBrowserFactory fileSystem)
        {
            return new IgnoredProjectsSettingsViewModel(provider, fileSystem, null);
        }

        private static Mock<IConfigurationService<IgnoredProjectsSettings>> ConfigProviderMock(IgnoredProjectsSettings initialSettings = null)
        {
            var mock = new Mock<IConfigurationService<IgnoredProjectsSettings>>();
            var defaultSettings = new IgnoredProjectsSettings();
            var currentSettings = initialSettings ?? new IgnoredProjectsSettings();

            mock.Setup(m => m.ReadDefaults()).Returns(defaultSettings);
            mock.Setup(m => m.Read()).Returns(currentSettings);

            return mock;
        }

        private static Mock<IFileSystemBrowserFactory> FileSystemBrowserMock(string newIgnoredFilename = null)
        {
            var dialogMock = FileDialogMock(newIgnoredFilename);
            return FileSystemBrowserMock(dialogMock.Object, newIgnoredFilename);
        }

        private static Mock<IFileSystemBrowserFactory> FileSystemBrowserMock(IOpenFileDialog fileDialog, string newIgnoredFilename = null)
        {
            var mock = new Mock<IFileSystemBrowserFactory>();
            mock.Setup(m => m.CreateOpenFileDialog()).Returns(fileDialog);
            return mock;
        }

        private static Mock<IOpenFileDialog> FileDialogMock(string newIgnoredFilename = null, bool dialogOk = true)
        {
            var mock = new Mock<IOpenFileDialog>();
            var dialogResult = !dialogOk || newIgnoredFilename == null
                ? DialogResult.Cancel
                : DialogResult.OK;

            mock.Setup(m => m.FileName).Returns(newIgnoredFilename);
            mock.Setup(m => m.ShowDialog()).Returns(dialogResult);
            
            return mock;
        }
    }
}