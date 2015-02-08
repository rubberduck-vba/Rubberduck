using System;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.Refactorings.ExtractMethod;
using Rubberduck.VBA;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class RefactorMenu : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addin;
        private readonly IRubberduckParser _parser;

        public RefactorMenu(VBE vbe, AddIn addin, IRubberduckParser parser)
        {
            _vbe = vbe;
            _addin = addin;
            _parser = parser;
        }

        private CommandBarButton _extractMethodButton;
        public CommandBarButton ExtractMethodButton { get { return _extractMethodButton; } }
        
        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "&Refactor";

            _extractMethodButton = AddMenuButton(menu);
            _extractMethodButton.Caption = "Extract &Method";
            _extractMethodButton.Click += OnExtractMethodButtonClick;

        }

        private void OnExtractMethodButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_vbe.ActiveCodePane == null)
            {
                return;
            }

            var selection = _vbe.ActiveCodePane.GetSelection();
            if (selection.StartLine <= _vbe.ActiveCodePane.CodeModule.CountOfDeclarationLines)
            {
                return;
            }

            vbext_ProcKind startKind;
            var startScope = _vbe.ActiveCodePane.CodeModule.get_ProcOfLine(selection.StartLine, out startKind);
            vbext_ProcKind endKind;
            var endScope = _vbe.ActiveCodePane.CodeModule.get_ProcOfLine(selection.EndLine, out endKind);

            if (startScope != endScope)
            {
                return;
            }

            // if method is a property, GetProcedure(name) can return up to 3 members:
            var method = _parser.Parse(_vbe.ActiveCodePane.CodeModule.Lines())
                                .GetProcedure(startScope)
                                .SingleOrDefault(proc => proc.GetSelection().Contains(selection));

            if (method == null)
            {
                return;
            }

            var view = new ExtractMethodDialog();
            var presenter = new ExtractMethodPresenter(view, method, selection);
            presenter.Show();
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(MsoControlType.msoControlButton) as CommandBarButton;
        }

        public void Dispose()
        {
            
        }
    }
}
