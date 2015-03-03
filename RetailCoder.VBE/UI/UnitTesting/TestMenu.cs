using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
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

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            menu.Caption = "&Test Explorer";
            menu.Click += OnTestExplorerButtonClick;
        }

        private void OnTestExplorerButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _presenter.Show();
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
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

            _disposed = true;
            base.Dispose(disposing);
        }
    }
}
