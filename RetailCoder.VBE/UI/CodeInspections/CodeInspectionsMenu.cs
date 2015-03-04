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
        private readonly CodeInspectionsDockablePresenter _presenter; //if presenter goes out of scope, so does its toolwindow Issue #169

        public CodeInspectionsMenu(VBE vbe, AddIn addIn, CodeInspectionsDockablePresenter presenter)
            :base(vbe, addIn)
        {
            _presenter = presenter;
        }

        public void Initialize(CommandBarPopup parentMenu)
        {
            AddButton(parentMenu, "Code &Inspections", false, new CommandBarButtonClickEvent(OnCodeInspectionsButtonClick));
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            _presenter.Show();
        }
    }
}
