using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.UI.SourceControl;
using Rubberduck.SourceControl;
using Moq;
using Rubberduck.Settings;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class SettingsPresenterTests
    {
        private const string Name = "Chris McClellan";
        private const string Email = "ckuhn203@gmail";
        private const string RepoLocation = @"C:\Users\Christopher\Documents";

        private const string OtherName= "King Lear";
        private const string OtherEmail = "king.lear@yahoo.com";
        private const string OtherRepoLocation = @"C:\Users\KingLear\Documents";

        private Mock<ISettingsView> _view;
        private Mock<IConfigurationService<SourceControlConfiguration>> _configService;
        private SourceControlConfiguration _config;

        [TestInitialize]
        public void Initialize()
        {
            _view = new Mock<ISettingsView>();
            _view.SetupProperty(v => v.UserName, string.Empty);
            _view.SetupProperty(v => v.EmailAddress, string.Empty);
            _view.SetupProperty(v => v.DefaultRepositoryLocation, string.Empty);

            _config = new SourceControlConfiguration(Name, Email, RepoLocation, new List<Repository>());

            _configService = new Mock<IConfigurationService<SourceControlConfiguration>>();
            _configService.Setup(s => s.LoadConfiguration()).Returns(_config);
        }

        [TestMethod]
        public void ViewIsPopulatedOnRefresh()
        {
            //arrange
            var presenter = new SettingsPresenter(_view.Object, _configService.Object);

            //act
            presenter.RefreshView();

            //assert
            Assert.AreEqual(Name, _view.Object.UserName, "Name");
            Assert.AreEqual(Email, _view.Object.EmailAddress, "Email");
            Assert.AreEqual(RepoLocation, _view.Object.DefaultRepositoryLocation, "Default Repo Location");
        }

        [TestMethod]
        public void ConfigIsPopulatedFromViewOnSave()
        {
            //arrange
            var presenter = new SettingsPresenter(_view.Object, _configService.Object);

            //simulate user input
            _view.Object.UserName = OtherName;
            _view.Object.EmailAddress = OtherEmail;
            _view.Object.DefaultRepositoryLocation = OtherRepoLocation;

            //simulate Update button click
            _view.Raise(v => v.Save += null, EventArgs.Empty);

            //assert
            Assert.AreEqual(OtherName, _config.UserName, "Name");
            Assert.AreEqual(OtherEmail, _config.EmailAddress, "Email");
            Assert.AreEqual(OtherRepoLocation, _config.DefaultRepositoryLocation, "Default Repo Location");
        }

        [TestMethod]
        public void ConfigIsSavedOnSave()
        {
            //arrange
            var presenter = new SettingsPresenter(_view.Object, _configService.Object);

            //act
            //simulate Update button click
            _view.Raise(v => v.Save += null, EventArgs.Empty);

            //assert
            _configService.Verify(s => s.SaveConfiguration(_config));
        }

        [TestMethod]
        public void ChangesToViewAreRevertedOnCancel()
        {
            //arrange
            var presenter = new SettingsPresenter(_view.Object, _configService.Object);

            //simulate user input
            _view.Object.UserName = OtherName;
            _view.Object.EmailAddress = OtherEmail;
            _view.Object.DefaultRepositoryLocation = OtherRepoLocation;

            //act
            //simulate Cancel button click
            _view.Raise(v => v.Cancel += null, EventArgs.Empty);

            //assert
            Assert.AreEqual(Name, _view.Object.UserName, "Name");
            Assert.AreEqual(Email, _view.Object.EmailAddress, "Email");
            Assert.AreEqual(RepoLocation, _view.Object.DefaultRepositoryLocation, "Default Repo Location");     
        }
    }
}
