using System.Windows.Forms.VisualStyles;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Properties;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI.CodeInspections
{
    public class CodeInspectionsMenu : Menu
    {
        private CommandBarButton _codeInspectionsButton;
        private readonly CodeInspectionsWindow _window;
        private readonly CodeInspectionsDockablePresenter _presenter; //if presenter goes out of scope, so does it's toolwindow Issue #169

        public CodeInspectionsMenu(VBE vbe, AddIn addIn, CodeInspectionsWindow view, CodeInspectionsDockablePresenter presenter)
            :base(vbe, addIn)
        {
            _window = view;
            _presenter = presenter;
        }

        public void Initialize(CommandBarPopup parentMenu)
        {
            _codeInspectionsButton = AddButton(parentMenu, RubberduckUI.RubberduckMenu_CodeInspections, false, new CommandBarButtonClickEvent(OnCodeInspectionsButtonClick));
        }

        public void Initialize(CommandBarControls menu, int beforeIndex, string caption)
        {
            _codeInspectionsButton = menu.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeIndex) as CommandBarButton;
            _codeInspectionsButton.Caption = caption;
            _codeInspectionsButton.Click += OnCodeInspectionsButtonClick;
        }

        public void Inspect()
        {
            _presenter.Show();
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            Inspect();
        }

        bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing && _window != null)
            {
                _window.Dispose();
            }

            _codeInspectionsButton.Click -= OnCodeInspectionsButtonClick;

            _disposed = true;
            base.Dispose(disposing);
        }
    }
}
