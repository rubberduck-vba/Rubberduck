using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.VBA;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class CodeInspectionsMenu
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IEnumerable<IInspection> _inspections;
        private readonly Parser _parser;

        public CodeInspectionsMenu(VBE vbe, AddIn addin, Parser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
            _inspections = inspections;
        }

        private CommandBarButton _codeInspectionsButton;

        public void Initialize(CommandBarControls menuControls)
        {
            _codeInspectionsButton = menuControls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            Debug.Assert(_codeInspectionsButton != null);

            _codeInspectionsButton.Caption = "Code &Inspections";

            _codeInspectionsButton.Click += OnCodeInspectionsButtonClick;
        }

        private void OnCodeInspectionsButtonClick(CommandBarButton ctrl, ref bool canceldefault)
        {
            var presenter = new CodeInspectionsDockablePresenter(_parser, _inspections, _vbe, _addin);
            presenter.Show();
        }
    }
}
