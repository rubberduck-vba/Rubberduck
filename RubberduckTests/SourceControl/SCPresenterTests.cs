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
using RubberduckTests.Mocks;

namespace RubberduckTests.SourceControl
{
    [TestClass]
    public class ScPresenterTests
    {
        private Mock<VBE> _vbe;
        private Windows _windows;
        private Mock<AddIn> _addIn;
        private Mock<Window> _window;
        private object _toolWindow;

        private Mock<ISourceControlView> _view;

        private Mock<IChangesPresenter> _changesPresenter;
        private Mock<IBranchesPresenter> _branchesPresenter;
        private Mock<ISettingsPresenter> _settingsPresenter;

        private Mock<IConfigurationService<SourceControlConfiguration>> _configService;

        private Mock<IFolderBrowserFactory> _folderBrowserFactory;
        private Mock<IFolderBrowser> _folderBrowser;
        
        private Mock<ISourceControlProviderFactory> _providerFactory;
        private Mock<ISourceControlProvider> _provider;

        [TestInitialize]
        public void InitializeMocks()
        {
            _window = new Mock<Window>();
            _window.SetupProperty(w => w.Visible, false);
            _window.SetupGet(w => w.LinkedWindows).Returns((LinkedWindows)null);
            _window.SetupProperty(w => w.Height);
            _window.SetupProperty(w => w.Width);

            _windows = new MockWindowsCollection(_window.Object);

            _vbe = new Mock<VBE>();
            _vbe.Setup(v => v.Windows).Returns(_windows);

            _addIn = new Mock<AddIn>();

            _view = new Mock<ISourceControlView>();
            _changesPresenter = new Mock<IChangesPresenter>();
            _branchesPresenter = new Mock<IBranchesPresenter>();
            _settingsPresenter = new Mock<ISettingsPresenter>();

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
            _providerFactory.Setup(f => f.CreateProvider(It.IsAny<VBProject>(), It.IsAny<IRepository>()))
                .Returns(_provider.Object);
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
       
            var branchesPresenter = new BranchesPresenter(branchesView.Object, new Mock<ICreateBranchView>().Object, new Mock<IMergeView>().Object);

            var provider = new Mock<ISourceControlProvider>();
            provider.Setup(git => git.Checkout(It.IsAny<string>()));
            provider.SetupGet(git => git.CurrentBranch)
                .Returns(new Branch("dev", "/ref/head/dev", false, true));

            branchesPresenter.Provider = provider.Object;
            changesPresenter.Provider = provider.Object;

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                                        _view.Object, changesPresenter, branchesPresenter,
                                                        _settingsPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object, 
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                                        _settingsPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object, 
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                                        _settingsPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object);

            //act
                _view.Raise(v => v.RefreshData += null, new EventArgs());

            //assert
            _changesPresenter.Verify(c => c.Refresh(), Times.Once);
        }

        [TestMethod]
        public void StatusIsOfflineWhenNoRepoIsFoundInConfig()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration()).Returns(new SourceControlConfiguration());

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                                        _settingsPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object);

            SetupValidVbProject();

            //act
            presenter.RefreshChildren();

            //assert
            Assert.AreEqual("Offline", _view.Object.Status);
            _changesPresenter.Verify(c => c.Refresh(), Times.Never);
            _branchesPresenter.Verify(b => b.RefreshView(), Times.Never);
        }

        [TestMethod]
        public void StatusIsOfflineWhenRepoListIsEmpty()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(new SourceControlConfiguration() { Repositories = new List<Repository>() });

            SetupValidVbProject();

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                                        _settingsPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            presenter.RefreshChildren();

            //assert
            Assert.AreEqual("Offline", _view.Object.Status);
            _changesPresenter.Verify(c => c.Refresh(), Times.Never);
            _branchesPresenter.Verify(b => b.RefreshView(), Times.Never);
        }

        [TestMethod]
        public void StatusIsOfflineIfNoMatchingRepoExists()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            var project = new Mock<VBProject>().SetupProperty(p => p.Name, "FooBar");
            _vbe.SetupProperty(vbe => vbe.ActiveVBProject, project.Object);

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                                        _settingsPresenter.Object,
                                                        _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            presenter.RefreshChildren();

            //assert
            Assert.AreEqual("Offline", _view.Object.Status);
            _changesPresenter.Verify(c => c.Refresh(), Times.Never);
            _branchesPresenter.Verify(b => b.RefreshView(), Times.Never);
        }

        [TestMethod]
        public void StatusIsOfflineWhenMultipleReposAreFound()
        {
            //arrange
            var config = GetDummyConfig();
            config.Repositories.Add(new Repository() { Name = dummyRepoName });

            _configService.Setup(c => c.LoadConfiguration())
                            .Returns(config);

            SetupValidVbProject();

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                            _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                            _settingsPresenter.Object,
                                            _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            presenter.RefreshChildren();

            //assert
            Assert.AreEqual("Offline", _view.Object.Status);
            _changesPresenter.Verify(c => c.Refresh(), Times.Never);
            _branchesPresenter.Verify(b => b.RefreshView(), Times.Never);

        }

        [TestMethod]
        public void StatusIsOnlineWhenRepoIsFound()
        {
            //arrange 
            _configService.Setup(c => c.LoadConfiguration())
                            .Returns(GetDummyConfig());

            SetupValidVbProject();

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                            _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                            _settingsPresenter.Object,
                                            _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                            _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                            _settingsPresenter.Object,
                                            _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            presenter.RefreshChildren();

            //assert
            Assert.IsNotNull(_changesPresenter.Object.Provider);
            Assert.IsNotNull(_branchesPresenter.Object.Provider);
        }

        [TestMethod]
        public void InitRepository_WhenUserCancels_RepoIsNotAddedToConfig()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            _view.Raise(v => v.InitializeNewRepository +=null, EventArgs.Empty);

            //assert
            _configService.Verify(c => c.SaveConfiguration(It.IsAny<SourceControlConfiguration>()), Times.Never);
        }

        [TestMethod]
        public void InitRepository_WhenUserCancels_RepoIsNotCreated()
        {
            //arrange
            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.Cancel);

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            _view.Raise(v => v.InitializeNewRepository += null, EventArgs.Empty);

            //assert
            _provider.Verify(git => git.InitVBAProject(It.IsAny<string>()),Times.Never);
        }

        [TestMethod]
        public void InitRepository_WhenUserConfirms_RepoIsAddedToConfig()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            _folderBrowser.Setup(b => b.ShowDialog()).Returns(DialogResult.OK);
            _folderBrowser.SetupProperty(b => b.SelectedPath, @"C:\path\to\repo\");

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            Assert.IsNotNull(_settingsPresenter.Object.Provider, "_settingsPresenter.Provider was null");
            Assert.IsNotNull(_branchesPresenter, "_branchesPresenter.Provider was null");
            Assert.IsNotNull(_changesPresenter, "_changesPresenter.Provider was null");
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

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object,
                                _view.Object, _changesPresenter.Object, _branchesPresenter.Object,
                                _settingsPresenter.Object,
                                _folderBrowserFactory.Object, _providerFactory.Object);

            //act
            _view.Raise(v => v.OpenWorkingDirectory += null, EventArgs.Empty);

            //assert
            Assert.IsNotNull(_settingsPresenter.Object.Provider, "_settingsPresenter.Provider was null");
            Assert.IsNotNull(_branchesPresenter, "_branchesPresenter.Provider was null");
            Assert.IsNotNull(_changesPresenter, "_changesPresenter.Provider was null");
        }

        private void SetupValidVbProject()
        {
            var project = new Mock<VBProject>().SetupProperty(p => p.Name, dummyRepoName);
            _vbe.SetupProperty(vbe => vbe.ActiveVBProject, project.Object);
        }

        private const string dummyRepoName = "SourceControlTest";

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
                           dummyRepoName,
                           @"C:\Users\Christopher\Documents\SourceControlTest",
                           @"https://github.com/ckuhn203/SourceControlTest.git"
                       );
        }
    }
}
