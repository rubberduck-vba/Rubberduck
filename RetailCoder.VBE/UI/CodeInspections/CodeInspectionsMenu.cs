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
        private readonly IEnumerable<IInspection> _inspections;
        private readonly IRubberduckParser _parser;
        private readonly CodeInspectionsWindow _window;
        private CodeInspectionsDockablePresenter _presenter; //if presenter goes out of scope, so does it's toolwindow Issue #169

        public CodeInspectionsMenu(VBE vbe, AddIn addin, IRubberduckParser parser, IEnumerable<IInspection> inspections)
            :base(vbe, addin)
        {
            _parser = parser;
            _inspections = inspections;
            //todo: inject dependencies;
            _window = new CodeInspectionsWindow();
            _presenter = new CodeInspectionsDockablePresenter(_parser, _inspections, this.IDE, this.addInInstance, _window);
        }

        private CommandBarButton _codeInspectionsButton;

        public void Initialize(CommandBarPopup parentMenu)
        {
            _codeInspectionsButton = AddButton(parentMenu, "Code &Inspections", false, new CommandBarButtonClickEvent(OnCodeInspectionsButtonClick));
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            _presenter.Show();
        }
    }
}
