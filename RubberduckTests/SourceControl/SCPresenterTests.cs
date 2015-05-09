using System;
using System.ComponentModel.Design;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using  Rubberduck.SourceControl;
using  Rubberduck.UI.SourceControl;
using  Moq;
using RubberduckTests.Mocks;
using Rubberduck.Config;
using System.Collections.Generic;

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
        private Mock<IConfigurationService<SourceControlConfiguration>> _configService;

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

            _configService = new Mock<IConfigurationService<SourceControlConfiguration>>();

            _view.SetupProperty(v => v.Status, string.Empty);
            
        }

        [TestMethod]
        public void BranchesRefreshOnRefreshEvent()
        {
            //arrange
            _configService.Setup(c => c.LoadConfiguration())
                .Returns(GetDummyConfig());

            SetupValidVbProject();

            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object, 
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                            _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                            _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

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
                                            _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

            //act
            presenter.RefreshChildren();

            //assert
            Assert.IsNotNull(_changesPresenter.Object.Provider);
            Assert.IsNotNull(_branchesPresenter.Object.Provider);
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
                            new Repository 
                            (
                                dummyRepoName,
                                @"C:\Users\Christopher\Documents\SourceControlTest",
                                @"https://github.com/ckuhn203/SourceControlTest.git"
                            )
                        }
                    };
        }
    }
}
