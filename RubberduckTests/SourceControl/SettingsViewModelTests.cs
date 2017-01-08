﻿using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.SettingsProvider;
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
        private const string CommandPromptLocation = "cmd.exe";

        private const string OtherName = "King Lear";
        private const string OtherEmail = "king.lear@yahoo.com";
        private const string OtherRepoLocation = @"C:\Users\KingLear\Documents";
        private const string OtherCommandPromptLocation = @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe";

        private Mock<IConfigProvider<SourceControlSettings>> _configService;
        private SourceControlSettings _config;

        private Mock<IFolderBrowserFactory> _folderBrowserFactory;
        private Mock<IFolderBrowser> _folderBrowser;

        private Mock<IOpenFileDialog> _openFileDialog;

        [TestInitialize]
        public void Initialize()
        {
            _config = new SourceControlSettings(Name, Email, RepoLocation, new List<Repository>(), CommandPromptLocation);

            _configService = new Mock<IConfigProvider<SourceControlSettings>>();
            _configService.Setup(s => s.Create()).Returns(_config);

            _folderBrowser = new Mock<IFolderBrowser>();
            _folderBrowserFactory = new Mock<IFolderBrowserFactory>();
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>())).Returns(_folderBrowser.Object);
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>(), false)).Returns(_folderBrowser.Object);

            _openFileDialog = new Mock<IOpenFileDialog>();
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ViewIsPopulatedOnRefresh()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object);

            //act
            vm.RefreshView();

            //assert
            Assert.AreEqual(Name, vm.UserName, "Name");
            Assert.AreEqual(Email, vm.EmailAddress, "Email");
            Assert.AreEqual(RepoLocation, vm.DefaultRepositoryLocation, "Default Repo Location");
            Assert.AreEqual(CommandPromptLocation, vm.CommandPromptLocation, "Command Prompt Location");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ConfigIsPopulatedFromViewOnSave()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object);

            //simulate user input
            vm.UserName = OtherName;
            vm.EmailAddress = OtherEmail;
            vm.DefaultRepositoryLocation = OtherRepoLocation;
            vm.CommandPromptLocation = OtherCommandPromptLocation;

            //simulate Update button click
            vm.UpdateSettingsCommand.Execute(null);

            //assert
            Assert.AreEqual(OtherName, _config.UserName, "Name");
            Assert.AreEqual(OtherEmail, _config.EmailAddress, "Email");
            Assert.AreEqual(OtherRepoLocation, _config.DefaultRepositoryLocation, "Default Repo Location");
            Assert.AreEqual(OtherCommandPromptLocation, _config.CommandPromptLocation, "Command Prompt Location");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ConfigIsSavedOnSave()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object);

            //act
            //simulate Update button click
            vm.UpdateSettingsCommand.Execute(null);

            //assert
            _configService.Verify(s => s.Save(_config));
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void ChangesToViewAreRevertedOnCancel()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object);

            //simulate user input
            vm.UserName = OtherName;
            vm.EmailAddress = OtherEmail;
            vm.DefaultRepositoryLocation = OtherRepoLocation;
            vm.DefaultRepositoryLocation = OtherCommandPromptLocation;

            //act
            //simulate Cancel button click
            vm.CancelSettingsChangesCommand.Execute(null);

            //assert
            Assert.AreEqual(Name, vm.UserName, "Name");
            Assert.AreEqual(Email, vm.EmailAddress, "Email");
            Assert.AreEqual(RepoLocation, vm.DefaultRepositoryLocation, "Default Repo Location");
            Assert.AreEqual(CommandPromptLocation, vm.CommandPromptLocation, "Command Prompt Location");
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnBrowseDefaultRepoLocation_WhenUserConfirms_ViewMatchesSelectedPath()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object)
            {
                DefaultRepositoryLocation = RepoLocation
            };
            _folderBrowser.Object.SelectedPath = OtherRepoLocation;
            _folderBrowser.Setup(f => f.ShowDialog()).Returns(DialogResult.OK);
            
            //act
            vm.ShowDefaultRepoFolderPickerCommand.Execute(null);

            //assert
            Assert.AreEqual(_folderBrowser.Object.SelectedPath, vm.DefaultRepositoryLocation);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnBrowseDefaultRepoLocation_WhenUserCancels_ViewRemainsUnchanged()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object)
            {
                DefaultRepositoryLocation = RepoLocation
            };
            _folderBrowser.Object.SelectedPath = OtherRepoLocation;
            _folderBrowser.Setup(f => f.ShowDialog()).Returns(DialogResult.Cancel);

            //act
            vm.ShowDefaultRepoFolderPickerCommand.Execute(null);

            //assert
            Assert.AreEqual(RepoLocation, vm.DefaultRepositoryLocation);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnBrowseCommandPromptLocation_WhenUserConfirms_ViewMatchesSelectedPath()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object)
            {
                CommandPromptLocation = CommandPromptLocation
            };
            _openFileDialog.Setup(o => o.FileName).Returns(OtherCommandPromptLocation);
            _openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.OK);

            //act
            vm.ShowCommandPromptExePickerCommand.Execute(null);

            //assert
            Assert.AreEqual(_openFileDialog.Object.FileName, vm.CommandPromptLocation);
        }

        [TestCategory("SourceControl")]
        [TestMethod]
        public void OnBrowseCommandPromptLocation_WhenUserCancels_ViewRemainsUnchanged()
        {
            //arrange
            var vm = new SettingsViewViewModel(_configService.Object, _folderBrowserFactory.Object, _openFileDialog.Object)
            {
                CommandPromptLocation = CommandPromptLocation
            };
            _openFileDialog.Setup(o => o.FileName).Returns(OtherCommandPromptLocation);
            _openFileDialog.Setup(o => o.ShowDialog()).Returns(DialogResult.Cancel);

            //act
            vm.ShowDefaultRepoFolderPickerCommand.Execute(null);

            //assert
            Assert.AreEqual(CommandPromptLocation, vm.CommandPromptLocation);
        }
    }
}
