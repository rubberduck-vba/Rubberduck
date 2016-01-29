using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.SourceControl;
using Rubberduck.UI;
using Rubberduck.UI.SourceControl;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

// ReSharper disable UnusedVariable
// Resharper thinks the presenter and toolWindow aren't used, but I promise they are.

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class ScPresenterTests
    {
        private Mock<VBE> _vbe;
        private MockWindowsCollection _windows;
        private Mock<AddIn> _addIn;
        private Mock<Window> _window;
#pragma warning disable 169
        private object _toolWindow;
#pragma warning restore 169

        private Mock<ISourceControlView> _view;

        private Mock<IChangesPresenter> _changesPresenter;
        private Mock<IBranchesPresenter> _branchesPresenter;
        private Mock<ISettingsPresenter> _settingsPresenter;
        private Mock<IUnsyncedCommitsPresenter> _unsyncedPresenter;

        private Mock<IConfigurationService<SourceControlConfiguration>> _configService;

        private Mock<IFolderBrowserFactory> _folderBrowserFactory;
        private Mock<IFolderBrowser> _folderBrowser;

        private Mock<ISourceControlProviderFactory> _providerFactory;
        private Mock<ISourceControlProvider> _provider;

        private Mock<IFailedMessageView> _failedActionView;
        private Mock<ILoginView> _loginView;
        private Mock<ICloneRepositoryView> _cloneRepo;

        [TestInitialize]
        public void InitializeMocks()
        {
            _window = Mocks.MockFactory.CreateWindowMock();
            _windows = new MockWindowsCollection(new List<Window> { _window.Object });
            _vbe = Mocks.MockFactory.CreateVbeMock(_windows);

            _addIn = new Mock<AddIn>();

            _view = new Mock<ISourceControlView>();
            _changesPresenter = new Mock<IChangesPresenter>();
            _branchesPresenter = new Mock<IBranchesPresenter>();
            _settingsPresenter = new Mock<ISettingsPresenter>();
            _unsyncedPresenter = new Mock<IUnsyncedCommitsPresenter>();

            _failedActionView = new Mock<IFailedMessageView>();
            _loginView = new Mock<ILoginView>();
            _cloneRepo = new Mock<ICloneRepositoryView>();

            _configService = new Mock<IConfigurationService<SourceControlConfiguration>>();

            _view.SetupProperty(v => v.Status, string.Empty);

            _folderBrowser = new Mock<IFolderBrowser>();
            _folderBrowserFactory = new Mock<IFolderBrowserFactory>();
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>())).Returns(_folderBrowser.Object);
            _folderBrowserFactory.Setup(f => f.CreateFolderBrowser(It.IsAny<string>(), false)).Returns(_folderBrowser.Object);

            _provider = new Mock<ISourceControlProvider>();
            _provider.Setup(git => git.InitVBAProject(It.IsAny<string>())).Returns(GetDummyRepo());

            _providerFactory = new Mock<ISourceControlProviderFactory>();
            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>()))
                .Returns(_provider.Object);
            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<ICodePaneWrapperFactory>()))
                .Returns(_provider.Object);
        }

        private SourceControlPresenter CreatePresenter()
        {
            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                _settingsPresenter.Object, _unsyncedPresenter.Object,
                _folderBrowserFactory.Object, _providerFactory.Object,
                _failedActionView.Object, _loginView.Object, _cloneRepo.Object, new CodePaneWrapperFactory());
            return presenter;
        }

        private void SetupValidVbProject()
        {
            var project = new Mock<VBProject>().SetupProperty(p => p.Name, DummyRepoName);
            _vbe.SetupProperty(vbe => vbe.ActiveVBProject, project.Object);
        }

        private void VerifyOffline()
        {
            Assert.AreEqual("Offline", _view.Object.Status);
            _changesPresenter.Verify(c => c.RefreshView(), Times.Never);
            _branchesPresenter.Verify(b => b.RefreshView(), Times.Never);
        }

        private void VerifyChildPresentersHaveProviders()
        {
            Assert.IsNotNull(_settingsPresenter.Object.Provider, "_settingsPresenter.Provider was null");
            Assert.IsNotNull(_branchesPresenter.Object.Provider, "_branchesPresenter.Provider was null");
            Assert.IsNotNull(_changesPresenter.Object.Provider, "_changesPresenter.Provider was null");
            Assert.IsNotNull(_unsyncedPresenter.Object.Provider, "_unsyncedPresenter.Object.Provider was null");
        }

        [TestMethod]
        public void ChangesCurrentBranchRefreshesWhenBranchIsCheckedOut()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            var changesView = new Mock<IChangesView>();
            changesView.SetupProperty(v => v.CurrentBranch, "master");
            var changesPresenter = new ChangesPresenter(changesView.Object);

            var branchesView = new Mock<IBranchesView>();
            branchesView.SetupProperty(b => b.Current, "master");

            var branchesPresenter = new BranchesPresenter(branchesView.Object, new Mock<ICreateBranchView>().Object, new Mock<IDeleteBranchView>().Object, new Mock<IMergeView>().Object);

            var provider = new Mock<ISourceControlProvider>();
            provider.Setup(git => git.Checkout(It.IsAny<string>()));
            provider.SetupGet(git => git.CurrentBranch)
                .Returns(new Branch("dev", "/ref/head/dev", false, true));

            branchesPresenter.Provider = provider.Object;
            changesPresenter.Provider = provider.Object;

            //purposely createing a new presenter with specific child presenters
            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                                        _view.Object, changesPresenter, branchesPresenter,
                                                        _settingsPresenter.Object, _unsyncedPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object,
                                                        _failedActionView.Object, _loginView.Object, _cloneRepo.Object, new CodePaneWrapperFactory());

            //act
            branchesView.Object.Current = "dev";
            branchesView.Raise(b => b.SelectedBranchChanged += null, new EventArgs());

            //assert
            Assert.AreEqual("dev", changesView.Object.CurrentBranch);
        }

        [TestMethod]
        public void BranchesRefreshOnRefreshEvent()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.RefreshData += null, new EventArgs());

            //assert
            _branchesPresenter.Verify(b => b.RefreshView(), Times.Once());
        }

        [TestMethod]
        public void ChangesRefreshOnRefreshEvent()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.RefreshData += null, new EventArgs());

            //assert
            _changesPresenter.Verify(c => c.RefreshView(), Times.Once);
        }

        [TestMethod]
        public void StatusIsOfflineWhenNoRepoIsFoundInConfig()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration()).Returns(new SourceControlConfiguration());

            var presenter = CreatePresenter();

            SetupValidVbProject();

            //act
            presenter.RefreshChildren();

            //assert
            VerifyOffline();
        }

        [TestMethod]
        public void StatusIsOfflineWhenRepoListIsEmpty()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(new SourceControlConfiguration() { Repositories = new List<Repository>() });

            SetupValidVbProject();

            var presenter = CreatePresenter();

            //act
            presenter.RefreshChildren();

            //assert
            VerifyOffline();
        }

        [TestMethod]
        public void StatusIsOfflineIfNoMatchingRepoExists()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            var project = new Mock<VBProject>().SetupProperty(p => p.Name, "FooBar");
            _vbe.SetupProperty(vbe => vbe.ActiveVBProject, project.Object);

            var presenter = CreatePresenter();

            //act
            presenter.RefreshChildren();

            //assert
            VerifyOffline();
        }

        [TestMethod]
        public void StatusIsOfflineWhenMultipleReposAreFound()
        {
            //arrange
            var config = GetDummyConfig();
            config.Repositories.Add(new Repository() { Name = DummyRepoName });

            _configService.Setup(c => c.LoadConfiguration())
                            .Returns(config);

            SetupValidVbProject();

            var presenter = CreatePresenter();

            //act
            presenter.RefreshChildren();

            //assert
            VerifyOffline();

        }

        [TestMethod]
        public void StatusIsOnlineWhenRepoIsFound()
        {
            //arrange 
            _configService.Setup(c => c.LoadConfiguration())
                            .Returns(GetDummyConfig());

            SetupValidVbProject();

            var presenter = CreatePresenter();

            //act
            presenter.RefreshChildren();

            //assert
            Assert.AreEqual("Online", _view.Object.Status);
        }

        [TestMethod]
        public void ChildPresentersHaveValidProviderIfRepoIsFoundInConfig()
        {
            //arrange 
            _configService.Setup(c => c.LoadConfiguration())
                            .Returns(GetDummyConfig());

            SetupValidVbProject();

            _changesPresenter.SetupProperty(c => c.Provider);
            _branchesPresenter.SetupProperty(b => b.Provider);
            _settingsPresenter.SetupProperty(s => s.Provider);
            _unsyncedPresenter.SetupProperty(s => s.Provider);

            var presenter = CreatePresenter();

            //act
            presenter.RefreshChildren();

            //assert
            VerifyChildPresentersHaveProviders();
        }

        [TestMethod]
        public void InitRepository_WhenUserCancels_RepoIsNotAddedToConfig()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.InitializeNewRepository += null, EventArgs.Empty);

            //assert
            _configService.Verify(c => c.SaveConfiguration(It.IsAny<SourceControlConfiguration>()), Times.Never);
        }

        [TestMethod]
        public void InitRepository_WhenUserCancels_RepoIsNotCreated()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.InitializeNewRepository += null, EventArgs.Empty);

            //assert
            _provider.Verify(git => git.InitVBAProject(It.IsAny<string>()), Times.Never);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_RepoIsAddedToConfig()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.InitializeNewRepository += null, EventArgs.Empty);

            //assert
            _configService.Verify(c => c.SaveConfiguration(It.IsAny<SourceControlConfiguration>()), Times.Once);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_RepoIsInitalized()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.InitializeNewRepository += null, EventArgs.Empty);

            //assert
            _provider.Verify(git => git.InitVBAProject(It.IsAny<string>()), Times.Once);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserCancels_RepoIsNotAddedToConfig()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            _configService.Verify(c => c.SaveConfiguration(It.IsAny<SourceControlConfiguration>()), Times.Never);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserConfirms_RepoIsAddedToConfig()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            _configService.Verify(c => c.SaveConfiguration(It.IsAny<SourceControlConfiguration>()), Times.Once);
        }


        [TestMethod]
        public void InitRepository_WhenUserConfirms_StatusIsOnline()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.InitializeNewRepository += null, EventArgs.Empty);

            //assert
            Assert.AreEqual("Online", _view.Object.Status);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserConfirms_StatusIsOnline()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            Assert.AreEqual("Online", _view.Object.Status);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_ChildPresenterSourceControlProvidersAreSet()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            _settingsPresenter.SetupProperty(s => s.Provider);
            _branchesPresenter.SetupProperty(b => b.Provider);
            _changesPresenter.SetupProperty(c => c.Provider);
            _unsyncedPresenter.SetupProperty(u => u.Provider);

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            VerifyChildPresentersHaveProviders();
        }

        [TestMethod]
        public void OpenWorkingDir_WhenUserConfirms_ChildPresenterSourceControlProvidersAreSet()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            _settingsPresenter.SetupProperty(s => s.Provider);
            _branchesPresenter.SetupProperty(b => b.Provider);
            _changesPresenter.SetupProperty(c => c.Provider);
            _unsyncedPresenter.SetupProperty(u => u.Provider);

            var presenter = CreatePresenter();

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            VerifyChildPresentersHaveProviders();
        }

        [TestMethod]
        public void BranchesPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            _failedActionView.SetupProperty(v => v.Message, String.Empty);

            _view.SetupProperty(v => v.SecondaryPanelVisible, false);
            _view.SetupProperty(v => v.SecondaryPanel);
            _view.Object.SecondaryPanel = _failedActionView.Object;

            var presenter = CreatePresenter();

            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            //act
            _branchesPresenter.Raise(b => b.ActionFailed += null, new ActionFailedEventArgs(expectedTitle, expectedMessage));

            //assert
            Assert.IsTrue(_view.Object.SecondaryPanelVisible);

            var expected = expectedTitle + Environment.NewLine + expectedMessage;
            Assert.AreEqual(expected, _failedActionView.Object.Message);
        }

        [TestMethod]
        public void ChangesPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            _failedActionView.SetupProperty(v => v.Message, String.Empty);

            _view.SetupProperty(v => v.SecondaryPanelVisible, false);
            _view.SetupProperty(v => v.SecondaryPanel);
            _view.Object.SecondaryPanel = _failedActionView.Object;

            var presenter = CreatePresenter();

            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            //act
            _changesPresenter.Raise(c => c.ActionFailed += null, new ActionFailedEventArgs(expectedTitle, expectedMessage));

            //assert
            Assert.IsTrue(_view.Object.SecondaryPanelVisible);

            var expected = expectedTitle + Environment.NewLine + expectedMessage;
            Assert.AreEqual(expected, _failedActionView.Object.Message);
        }

        [TestMethod]
        public void SettingsPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            _failedActionView.SetupProperty(v => v.Message, String.Empty);

            _view.SetupProperty(v => v.SecondaryPanelVisible, false);
            _view.SetupProperty(v => v.SecondaryPanel);
            _view.Object.SecondaryPanel = _failedActionView.Object;

            var presenter = CreatePresenter();

            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            //act
            _settingsPresenter.Raise(s => s.ActionFailed += null, new ActionFailedEventArgs(expectedTitle, expectedMessage));

            //assert
            Assert.IsTrue(_view.Object.SecondaryPanelVisible);

            var expected = expectedTitle + Environment.NewLine + expectedMessage;
            Assert.AreEqual(expected, _failedActionView.Object.Message);
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenActionFailedEventIsRaised_MessageIsShown()
        {
            //arrange
            _failedActionView.SetupProperty(v => v.Message, String.Empty);

            _view.SetupProperty(v => v.SecondaryPanelVisible, false);
            _view.SetupProperty(v => v.SecondaryPanel);
            _view.Object.SecondaryPanel = _failedActionView.Object;

            var presenter = CreatePresenter();

            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            //act
            _unsyncedPresenter.Raise(u => u.ActionFailed += null, new ActionFailedEventArgs(expectedTitle, expectedMessage));

            //assert
            Assert.IsTrue(_view.Object.SecondaryPanelVisible);

            var expected = expectedTitle + Environment.NewLine + expectedMessage;
            Assert.AreEqual(expected, _failedActionView.Object.Message);
        }

        [TestMethod]
        public void OpenWorkingDir_WhenProviderCreationFails_MessageIsShown()
        {
            //arrange
            const string expectedTitle = "Some Action Failed.";
            const string expectedMessage = "Details about failure.";

            //arrange
            _failedActionView.SetupProperty(v => v.Message, String.Empty);

            _view.SetupProperty(v => v.SecondaryPanelVisible, false);
            _view.SetupProperty(v => v.SecondaryPanel);
            _view.Object.SecondaryPanel = _failedActionView.Object;

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<ICodePaneWrapperFactory>()))
                .Throws(new SourceControlException(expectedTitle,
                    new LibGit2Sharp.LibGit2SharpException(expectedMessage))
                    );

            SetupValidVbProject();

            var presenter = CreatePresenter();
            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            Assert.IsTrue(_view.Object.SecondaryPanelVisible);

            var expected = expectedTitle + Environment.NewLine + expectedMessage;
            Assert.AreEqual(expected, _failedActionView.Object.Message);
        }

        [TestMethod]
        public void ActionFailed_DismissingHidesMessage()
        {
            //arrange
            _view.SetupProperty(v => v.SecondaryPanelVisible, true);

            var presenter = CreatePresenter();

            //act
            _failedActionView.Raise(v => v.DismissSecondaryPanel += null, EventArgs.Empty);

            //assert
            Assert.IsFalse(_view.Object.SecondaryPanelVisible);
        }

        [TestMethod]
        public void UnsyncedPresenter_WhenNotAuthorized_LoginIsShown()
        {
            //arrange
            _failedActionView.SetupProperty(v => v.Message, String.Empty);

            _view.SetupProperty(v => v.SecondaryPanelVisible, false);
            _view.SetupProperty(v => v.SecondaryPanel);
            _view.Object.SecondaryPanel = _failedActionView.Object;

            var presenter = CreatePresenter();

            const string outerMessage = "Push Failed.";
            const string innerMessage = "Request failed with status code: 401";

            //act
            _unsyncedPresenter.Raise(u => u.ActionFailed += null, new ActionFailedEventArgs(outerMessage, innerMessage));

            //assert
            Assert.IsInstanceOfType(_view.Object.SecondaryPanel, typeof(ILoginView));
        }

        [TestMethod]
        public void UnsyncedPresenter_AfterLogin_NewPresenterIsCreatedWithCredentials()
        {
            //arrange
            const string username = "username";
            const string password = "password";

            _loginView.SetupProperty(v => v.Username, username);
            _loginView.SetupProperty(v => v.Password, password);

            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            var presenter = CreatePresenter();

            //act
            _loginView.Raise(v => v.Confirm += null, EventArgs.Empty);

            //assert
            _providerFactory.Verify(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>(), It.IsAny<SecureCredentials>(), It.IsAny<ICodePaneWrapperFactory>()));
        }

        private const string DummyRepoName = "SourceControlTest";

        private SourceControlConfiguration GetDummyConfig()
        {
            return new SourceControlConfiguration()
                    {
                        Repositories = new List<Repository>() 
                        { 
                            (Repository)GetDummyRepo()
                        }
                    };
        }

        private static IRepository GetDummyRepo()
        {
            return new Repository
                       (
                           DummyRepoName,
                           @"C:\Users\Christopher\Documents\SourceControlTest",
                           @"https://github.com/ckuhn203/SourceControlTest.git"
                       );
        }
    }
}
