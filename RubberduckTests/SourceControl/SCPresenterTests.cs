using System;
using System.ComponentModel.Design;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using  Rubberduck.SourceControl;
using  Rubberduck.UI.SourceControl;
using  Moq;
using RubberduckTests.Mocks;
using Rubberduck.Config;

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
        }

        [TestMethod]
        public void BranchesRefreshOnRefreshEvent()
        {
            //arrange
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
            var presenter = new SourceControlPresenter(_vbe.Object, _addIn.Object, _configService.Object, 
                                                        _view.Object, _changesPresenter.Object, _branchesPresenter.Object);

            //act
                _view.Raise(v => v.RefreshData += null, new EventArgs());

            //assert
            _changesPresenter.Verify(c => c.Refresh(), Times.Once);
        }
    }
}
