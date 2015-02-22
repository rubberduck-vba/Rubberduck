using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerMenu : Menu
    {
        private readonly IRubberduckParser _parser;
        private readonly CodeExplorerWindow _window;

        public CodeExplorerMenu(VBE vbe, AddIn addin, IRubberduckParser parser)
            :base(vbe, addin)
        {
            _parser = parser;
            _window = new CodeExplorerWindow();
        }

        private CommandBarButton _codeExplorerButton;

        public void Initialize(CommandBarPopup parentMenu)
        {
            _codeExplorerButton = AddButton(parentMenu, "&Code Explorer", true, new CommandBarButtonClickEvent(OnCodeExplorerButtonClick), 3039);
        }

        private void OnCodeExplorerButtonClick(CommandBarButton button, ref bool cancelDefault)
        {
            var presenter = new CodeExplorerDockablePresenter(_parser, this.IDE, this.addInInstance, _window);
            presenter.Show();
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
