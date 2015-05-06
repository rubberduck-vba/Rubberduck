using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerMenu : Menu
    {
        private CommandBarButton _codeExplorerButton;
        private readonly CodeExplorerWindow _window;
        private readonly CodeExplorerDockablePresenter _presenter; //if presenter goes out of scope, so does it's toolwindow Issue #169

        public CodeExplorerMenu(VBE vbe, AddIn addIn, CodeExplorerWindow view, CodeExplorerDockablePresenter presenter)
            :base(vbe, addIn)
        {
            _window = view;
            _presenter = presenter;
        }

        public void Initialize(CommandBarPopup parentMenu)
        {
            _codeExplorerButton = AddButton(parentMenu, "&Code Explorer", true, new CommandBarButtonClickEvent(OnCodeExplorerButtonClick), 3039);
        }

        private void OnCodeExplorerButtonClick(CommandBarButton button, ref bool cancelDefault)
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

            if (disposing && _window != null)
            {
                _window.Dispose();
            }

            disposed = true;
            base.Dispose(disposing);
        }
    }
}
