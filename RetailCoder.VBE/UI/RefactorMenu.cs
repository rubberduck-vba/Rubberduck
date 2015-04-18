using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.UI.IdentifierReferences;
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

            InitializeRefactorContextMenu();
        }

        private CommandBarPopup _refactorCodePaneContextMenu;
        public CommandBarPopup RefactorCodePaneContextMenu { get { return _refactorCodePaneContextMenu; } }

        private CommandBarButton _extractMethodContextButton;
        private CommandBarButton _renameContextButton;

        private void InitializeRefactorContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _refactorCodePaneContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true, Before:beforeItem) as CommandBarPopup;
            _refactorCodePaneContextMenu.BeginGroup = true;
            _refactorCodePaneContextMenu.Caption = "&Refactor";

            var extractMethodIcon = Resources.ExtractMethod_6786_32;
            extractMethodIcon.MakeTransparent(Color.White);
            _extractMethodContextButton = AddButton(_refactorCodePaneContextMenu, "Extract &Method", false, OnExtractMethodButtonClick, extractMethodIcon);
            _renameContextButton = AddButton(_refactorCodePaneContextMenu, "&Rename", false, OnRenameButtonClick);

            InitializeFindReferencesContextMenu(); //todo: untangle that mess...
        }

        private CommandBarButton _findAllReferencesContextMenu;
        private void InitializeFindReferencesContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findAllReferencesContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            _findAllReferencesContextMenu.Caption = "&Find all references...";
            _findAllReferencesContextMenu.Click += _findAllReferencesContextMenu_Click;
        }

        private void _findAllReferencesContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var selection = IDE.ActiveCodePane.GetSelection();
            var declarations = _parser.Parse(IDE.ActiveVBProject).Declarations;

            var target = declarations.Items
            .Where(item => item.DeclarationType != DeclarationType.ModuleOption)
            .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                  || IsSelectedReference(selection, item));

            if (target == null)
            {
                return;
            }

            FindAllReferences(target);
        }

        public void FindAllReferences(Declaration target)
        {
            var referenceCount = target.References.Count();

            if (referenceCount == 1)
            {
                // if there's only 1 reference, just jump to it:
                IdentifierReferencesListDockablePresenter.OnNavigateIdentifierReference(IDE, target.References.First());
            }
            else if (referenceCount > 1)
            {
                // if there's more than one reference, show the dockable reference navigation window:
                try
                {
                    ShowReferencesToolwindow(target);
                }
                catch (COMException)
                {
                    // the exception is related to the docked control host instance,
                    // trying again will work (I know, that's bad bad bad code)
                    ShowReferencesToolwindow(target);
                }
            }
        }

        private void ShowReferencesToolwindow(Declaration target)
        {
            // throws a COMException if toolwindow was already closed
            var window = new IdentifierReferencesListControl(target);
            var presenter = new IdentifierReferencesListDockablePresenter(IDE, AddIn, window, target);
            presenter.Show();
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.References.Any(r =>
                r.QualifiedModuleName == selection.QualifiedName &&
                r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
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
            Rename(selection);
        }

        public void Rename(QualifiedSelection selection)
        {
            using (var view = new RenameDialog())
            {
                var parseResult = _parser.Parse(IDE.ActiveVBProject);
                var presenter = new RenamePresenter(IDE, view, parseResult, selection);
                presenter.Show();
            }
        }

        public void Rename(Declaration target)
        {
            using (var view = new RenameDialog())
            {
                var parseResult = _parser.Parse(IDE.ActiveVBProject);
                var presenter = new RenamePresenter(IDE, view, parseResult, new QualifiedSelection(target.QualifiedName.QualifiedModuleName, target.Selection));
                presenter.Show(target);
            }
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(MsoControlType.msoControlButton) as CommandBarButton;
        }
    }
}
