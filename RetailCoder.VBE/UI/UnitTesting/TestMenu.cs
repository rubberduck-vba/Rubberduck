using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.VBIDEApi;
using Rubberduck.Properties;

namespace Rubberduck.UI.UnitTesting
{
    public class TestMenu : Menu
    {
        // 2743: play icon with stopwatch
        // 3039: module icon || 3119 || 621 || 589 || 472
        // 3170: class module icon

        private readonly TestExplorerWindow _view;
        private readonly TestExplorerDockablePresenter _presenter;
        public TestMenu(VBE vbe, AddIn addIn, TestExplorerWindow view, TestExplorerDockablePresenter presenter)
            : base(vbe, addIn)
        {
            _view = view;
            _presenter = presenter;
        }

        private CommandBarButton _runAllTestsButton;
        private CommandBarButton _windowsTestExplorerButton;

        public void Initialize(CommandBarControls menuControls)
        {
            _menuControls = menuControls;

            _menu = menuControls.Add(MsoControlType.msoControlPopup, null, null, null, true) as CommandBarPopup;
            _menu.Caption = RubberduckUI.RubberduckMenu_UnitTests;

            _windowsTestExplorerButton = AddButton(_menu, RubberduckUI.TestMenu_TextExplorer, false, OnTestExplorerButtonClick);
            SetButtonImage(_windowsTestExplorerButton, Resources.TestManager_8590_32, Resources.TestManager_8590_32_Mask);

            _runAllTestsButton = AddButton(_menu, RubberduckUI.TestMenu_RunAllTests, true, OnRunAllTestsButtonClick);
            SetButtonImage(_runAllTestsButton, Resources.AllLoadedTests_8644_24, Resources.AllLoadedTests_8644_24_Mask);
        }

        public void RunAllTests()
        {
            var cancelDefault = false;
            OnRunAllTestsButtonClick(null, ref cancelDefault);
        }

        void OnRunAllTestsButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _presenter.Show();
            _presenter.RunTests();
        }

        void OnTestExplorerButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _presenter.Show();
        }

        bool _disposed;
        private CommandBarPopup _menu;
        private CommandBarControls _menuControls;

        protected override void Dispose(bool disposing)
        {
            if (_disposed || !disposing)
            {
                return;
            }

            if (_view != null)
            {
                _view.Dispose();
            }

            _menuControls.Parent.FindControl(_menu.Type, _menu.Id, _menu.Tag, _menu.Visible).Delete();

            _runAllTestsButton.ClickEvent -= OnRunAllTestsButtonClick;
            _windowsTestExplorerButton.ClickEvent -= OnTestExplorerButtonClick;

            _disposed = true;
            base.Dispose(disposing);
        }
    }
}
