using System.Linq;
using Antlr4.Runtime;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Listeners;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.UI.Refactorings.ExtractMethod;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.VBA;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI
{
    public class RefactorMenu : Menu
    {
        private readonly IRubberduckParser _parser;

        public RefactorMenu(VBE vbe, AddIn addin, IRubberduckParser parser)
            : base(vbe, addin)
        {
            _parser = parser;
        }

        private CommandBarButton _extractMethodButton;
        private CommandBarButton _renameButton;

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(Type: MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "&Refactor";

            _extractMethodButton = AddButton(menu, "Extract &Method", false, OnExtractMethodButtonClick, Resources.ExtractMethod_6786_32);
            _renameButton = AddButton(menu, "&Rename", false, OnRenameButtonClick);
        }

        private void OnExtractMethodButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }
            var selection = IDE.ActiveCodePane.GetSelection();
            if (selection.Selection.StartLine <= IDE.ActiveCodePane.CodeModule.CountOfDeclarationLines)
            {
                return;
            }

            vbext_ProcKind startKind;
            var startScope = IDE.ActiveCodePane.CodeModule.get_ProcOfLine(selection.Selection.StartLine, out startKind);
            vbext_ProcKind endKind;
            var endScope = IDE.ActiveCodePane.CodeModule.get_ProcOfLine(selection.Selection.EndLine, out endKind);

            if (startScope != endScope)
            {
                return;
            }

            // if method is a property, GetProcedure(name) can return up to 3 members:
            var method = (_parser.Parse(IDE.ActiveVBProject).Declarations.Items
                                .SingleOrDefault(declaration => 
                                    (declaration.DeclarationType == DeclarationType.Procedure
                                    || declaration.DeclarationType == DeclarationType.Function
                                    || declaration.DeclarationType == DeclarationType.PropertyGet
                                    || declaration.DeclarationType == DeclarationType.PropertyLet
                                    || declaration.DeclarationType == DeclarationType.PropertySet) 
                                && declaration.Context.GetSelection().Contains(selection.Selection)));

            if (method == null)
            {
                return;
            }

            var view = new ExtractMethodDialog();
            var presenter = new ExtractMethodPresenter(IDE, view, method.Context, selection);
            presenter.Show();

            view.Dispose();
        }

        private void OnRenameButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }
            var selection = IDE.ActiveCodePane.GetSelection();
            var view = new RenameDialog();
            var presenter = new RenamePresenter(IDE, view, _parser.Parse(IDE.ActiveVBProject).Declarations, selection);
            presenter.Show();

            view.Dispose();
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(MsoControlType.msoControlButton) as CommandBarButton;
        }
    }
}
