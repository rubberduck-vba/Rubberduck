using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.SourceControl;
using Rubberduck.UI;
using Rubberduck.UI.SourceControl;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class SettingsViewModelTests
    {
        private const string Name = "Chris McClellan";
        private const string Email = "ckuhn203@gmail";
        private const string RepoLocation = @"C:\Users\Christopher\Documents";

        private const string OtherName = "King Lear";
        private const string OtherEmail = "king.lear@yahoo.com";
        private const string OtherRepoLocation = @"C:\Users\KingLear\Documents";

        private Mock<ISourceControlConfigProvider> _configService;
        private SourceControlSettings _config;

        private Mock<IFolderBrowserFactory> _folderBrowserFactory;
        private Mock<IFolderBrowser> _folderBrowser;

        [TestInitialize]
        public void Initialize()
        {
            _config = new SourceControlSettings(Name, Email, RepoLocation, new List<Repository>());

            _configService = new Mock<ISourceControlConfigProvider>();
            _configService.Setup(s => s.Create()).Returns(_config);

            _folderBrowser = new Mock<IFolderBrowser>();
            _folderBrowserFactory = new Mock<IFolderBrowserFactory>();
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>())).Returns(_folderBrowser.Object);
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>(), false)).Returns(_folderBrowser.Object);
        }

        [TestMethod]
        public void ViewIsPopulatedOnRefresh()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object);

            //act
            vm.RefreshView();

            //assert
            Assert.AreEqual(Name, vm.UserName, "Name");
            Assert.AreEqual(Email, vm.EmailAddress, "Email");
            Assert.AreEqual(RepoLocation, vm.DefaultRepositoryLocation, "Default Repo Location");
        }

        [TestMethod]
        public void ConfigIsPopulatedFromViewOnSave()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object);

            //simulate user input
            vm.UserName = OtherName;
            vm.EmailAddress = OtherEmail;
            vm.DefaultRepositoryLocation = OtherRepoLocation;

            //simulate Update button click
            vm.UpdateSettingsCommand.Execute(null);

            //assert
            Assert.AreEqual(OtherName, _config.UserName, "Name");
            Assert.AreEqual(OtherEmail, _config.EmailAddress, "Email");
            Assert.AreEqual(OtherRepoLocation, _config.DefaultRepositoryLocation, "Default Repo Location");
        }

        [TestMethod]
        public void ConfigIsSavedOnSave()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object);

            //act
            //simulate Update button click
            vm.UpdateSettingsCommand.Execute(null);

            //assert
            _configService.Verify(s => s.Save(_config));
        }

        [TestMethod]
        public void ChangesToViewAreRevertedOnCancel()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object);

            //simulate user input
            vm.UserName = OtherName;
            vm.EmailAddress = OtherEmail;
            vm.DefaultRepositoryLocation = OtherRepoLocation;

            //act
            //simulate Cancel button click
            vm.CancelSettingsChangesCommand.Execute(null);

            //assert
            Assert.AreEqual(Name, vm.UserName, "Name");
            Assert.AreEqual(Email, vm.EmailAddress, "Email");
            Assert.AreEqual(RepoLocation, vm.DefaultRepositoryLocation, "Default Repo Location");
        }

        [TestMethod]
        public void OnBrowseDefaultRepoLocation_WhenUserConfirms_ViewMatchesSelectedPath()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object)
            {
                DefaultRepositoryLocation = RepoLocation
            };
            _folderBrowser.Object.SelectedPath = OtherRepoLocation;
            _folderBrowser.Setup(f => f.ShowDialog()).Returns(DialogResult.OK);
            
            //act
            vm.ShowFilePickerCommand.Execute(null);

            //assert
            Assert.AreEqual(_folderBrowser.Object.SelectedPath, vm.DefaultRepositoryLocation);
        }

        [TestMethod]
        public void OnBrowserDefaultRepoLocation_WhenUserCancels_ViewRemainsUnchanged()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object)
            {
                DefaultRepositoryLocation = RepoLocation
            };
            _folderBrowser.Object.SelectedPath = OtherRepoLocation;
            _folderBrowser.Setup(f => f.ShowDialog()).Returns(DialogResult.Cancel);

            //act
            vm.ShowFilePickerCommand.Execute(null);

            //assert
            Assert.AreEqual(RepoLocation, vm.DefaultRepositoryLocation);
        }
    }
}
