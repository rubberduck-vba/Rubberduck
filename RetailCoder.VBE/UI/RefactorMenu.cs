using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Properties;
using Rubberduck.UI.Refactorings.ExtractMethod;
using Rubberduck.VBA;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class RefactorMenu : Menu, IDisposable
    {
        private readonly IRubberduckParser _parser;

        public RefactorMenu(VBE vbe, AddIn addin, IRubberduckParser parser)
            : base(vbe, addin)
        {
            _parser = parser;
        }

        private CommandBarButton _extractMethodButton;

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "&Refactor";

            _extractMethodButton = AddMenuButton(menu,"Extract &Method", Resources.ExtractMethod_6786_32);
            _extractMethodButton.Click += OnExtractMethodButtonClick;

        }

        private void OnExtractMethodButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }

            var selection = IDE.ActiveCodePane.GetSelection();
            if (selection.StartLine <= IDE.ActiveCodePane.CodeModule.CountOfDeclarationLines)
            {
                return;
            }

            vbext_ProcKind startKind;
            var startScope = IDE.ActiveCodePane.CodeModule.get_ProcOfLine(selection.StartLine, out startKind);
            vbext_ProcKind endKind;
            var endScope = IDE.ActiveCodePane.CodeModule.get_ProcOfLine(selection.EndLine, out endKind);

            if (startScope != endScope)
            {
                return;
            }

            // if method is a property, GetProcedure(name) can return up to 3 members:
            var method = ((IEnumerable<ParserRuleContext>) _parser.Parse(IDE.ActiveCodePane.CodeModule.Lines()).GetContexts<ProcedureNameListener, ParserRuleContext>(new ProcedureNameListener(startScope)))
                                .SingleOrDefault(proc => proc.GetSelection().Contains(selection));

            if (method == null)
            {
                return;
            }

            var view = new ExtractMethodDialog();
            var presenter = new ExtractMethodPresenter(IDE, view, method, selection);
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
