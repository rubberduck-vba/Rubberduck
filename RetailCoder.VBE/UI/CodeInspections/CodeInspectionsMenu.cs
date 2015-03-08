using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.VBA;
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
            _codeInspectionsButton = AddButton(parentMenu, "Code &Inspections", false, new CommandBarButtonClickEvent(OnCodeInspectionsButtonClick));
        }

        public void Inspect()
        {
            _presenter.Show();
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            Inspect();
        }
    }
}
