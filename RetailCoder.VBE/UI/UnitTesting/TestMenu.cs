using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Properties;
using Rubberduck.UnitTesting;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

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
            var menu = menuControls.Add(MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "Te&st";

            _windowsTestExplorerButton = AddButton(menu, "&Test Explorer", false, new CommandBarButtonClickEvent(OnTestExplorerButtonClick), Resources.TestManager_8590_32);
            _runAllTestsButton = AddButton(menu, "&Run All Tests", true, new CommandBarButtonClickEvent(OnRunAllTestsButtonClick), Resources.AllLoadedTests_8644_24);
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

        bool disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                if (_view != null)
                {
                    _view.Dispose();
                }

                if (_presenter != null)
                {
                    _presenter.Dispose();
                }
            }

            disposed = true;
            base.Dispose(disposing);
        }
    }
}
