using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.UI.FindSymbol;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.UI.Refactorings.ExtractMethod;
using Rubberduck.UI.Refactorings.Rename;
using Rubberduck.UI.Refactorings.RemoveParameter;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI
{
    public class RefactorMenu : Menu
    {
        private readonly IRubberduckParser _parser;
        private readonly IActiveCodePaneEditor _editor;

        private readonly SearchResultIconCache _iconCache;

        public RefactorMenu(VBE vbe, AddIn addin, IRubberduckParser parser, IActiveCodePaneEditor editor)
            : base(vbe, addin)
        {
            _parser = parser;
            _editor = editor;

            _iconCache = new SearchResultIconCache();
        }

        private CommandBarButton _extractMethodButton;
        private CommandBarButton _renameButton;
        private CommandBarButton _reorderParametersButton;
        private CommandBarButton _removeParameterButton;

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = "&Refactor";

            _extractMethodButton = AddButton(menu, "Extract &Method", false, OnExtractMethodButtonClick);
            SetButtonImage(_extractMethodButton, Resources.ExtractMethod_6786_32, Resources.ExtractMethod_6786_32_Mask);

            _renameButton = AddButton(menu, "&Rename", false, OnRenameButtonClick);
            
            _reorderParametersButton = AddButton(menu, "Reorder &Parameters", false, OnReorderParametersButtonClick, Resources.ReorderParameters_6780_32);
            SetButtonImage(_reorderParametersButton, Resources.ReorderParameters_6780_32, Resources.ReorderParameters_6780_32_Mask);

            _removeParameterButton = AddButton(menu, "Remo&ve Parameter", false, OnRemoveParameterButtonClick);
            //SetButtonImage(_removeParameterButton, Resources.RemoveParameters_6781_32_Mask);

            InitializeRefactorContextMenu();
        }

        private CommandBarPopup _refactorCodePaneContextMenu;
        public CommandBarPopup RefactorCodePaneContextMenu { get { return _refactorCodePaneContextMenu; } }

        private CommandBarButton _extractMethodContextButton;
        private CommandBarButton _renameContextButton;
        private CommandBarButton _reorderParametersContextButton;
        private CommandBarButton _removeParameterContextButton;

        private void InitializeRefactorContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _refactorCodePaneContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true, Before:beforeItem) as CommandBarPopup;
            _refactorCodePaneContextMenu.BeginGroup = true;
            _refactorCodePaneContextMenu.Caption = "&Refactor";

            _extractMethodContextButton = AddButton(_refactorCodePaneContextMenu, "Extract &Method", false, OnExtractMethodButtonClick);
            SetButtonImage(_extractMethodContextButton, Resources.ExtractMethod_6786_32, Resources.ExtractMethod_6786_32_Mask);

            _renameContextButton = AddButton(_refactorCodePaneContextMenu, "&Rename", false, OnRenameButtonClick);

            _reorderParametersContextButton = AddButton(_refactorCodePaneContextMenu, "Reorder &Parameters", false, OnReorderParametersButtonClick);
            SetButtonImage(_reorderParametersContextButton, Resources.ReorderParameters_6780_32, Resources.ReorderParameters_6780_32_Mask);

            _removeParameterContextButton = AddButton(_refactorCodePaneContextMenu, "Remo&ve Parameter", false, OnRemoveParameterButtonClick);
            //SetButtonImage(_removeParameterButton, Resources.RemoveParameters_6781_32_Mask);

            InitializeFindReferencesContextMenu(); //todo: untangle that mess...
            InitializeFindSymbolContextMenu();
        }

        private CommandBarButton _findAllReferencesContextMenu;
        private void InitializeFindReferencesContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findAllReferencesContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            _findAllReferencesContextMenu.Caption = "&Find all references...";
            _findAllReferencesContextMenu.Click += _findAllReferencesContextMenu_Click;
        }

        private CommandBarButton _findSymbolContextMenu;
        private void InitializeFindSymbolContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findSymbolContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            SetButtonImage(_findSymbolContextMenu, Resources.FindSymbol_6263_32, Resources.FindSymbol_6263_32_Mask);
            _findSymbolContextMenu.Caption = "Find &Symbol...";
            _findSymbolContextMenu.Click += FindSymbolContextMenuClick;
        }

        private void FindSymbolContextMenuClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var declarations = _parser.Parse(IDE.ActiveVBProject, this).Declarations;
            var vm = new FindSymbolViewModel(declarations.Items.Where(item => !item.IsBuiltIn), _iconCache);
            vm.Navigate += vm_Navigate;
            using (var view = new FindSymbolDialog(vm))
            {
                view.ShowDialog();
            }

            vm.Navigate -= vm_Navigate;
        }

        private void vm_Navigate(object sender, NavigateCodeEventArgs e)
        {
            if (e.QualifiedName.Component == null)
            {
                return;
            }

            try
            {
                e.QualifiedName.Component.CodeModule.CodePane.SetSelection(e.Selection);
            }
            catch (COMException)
            {
            }
        }

        private void _findAllReferencesContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var selection = IDE.ActiveCodePane.GetSelection();
            var declarations = _parser.Parse(IDE.ActiveVBProject, this).Declarations;

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
            else
            {
                var message = string.Format(RubberduckUI.AllReferences_NoneFound, target.IdentifierName);
                var caption = string.Format(RubberduckUI.AllReferences_Caption, target.IdentifierName);
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            var selectedScope = _editor.GetSelectedProcedureScope(selection.Selection);           
            if (string.IsNullOrEmpty(selectedScope))
            {
                return;
            }

            var declarations = _parser.Parse(IDE.ActiveVBProject, this).Declarations;

            // if method is a property, GetProcedure(name) can return up to 3 members:
            var target = (declarations.Items
                                .SingleOrDefault(declaration => declaration.Scope == selectedScope &&
                                    (declaration.DeclarationType == DeclarationType.Procedure
                                    || declaration.DeclarationType == DeclarationType.Function
                                    || declaration.DeclarationType == DeclarationType.PropertyGet
                                    || declaration.DeclarationType == DeclarationType.PropertyLet
                                    || declaration.DeclarationType == DeclarationType.PropertySet)));

            if (target == null)
            {
                return;
            }

            var view = new ExtractMethodDialog();
            var presenter = new ExtractMethodPresenter(_editor, view, target, selection, declarations);
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

        private void OnReorderParametersButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }
            var selection = IDE.ActiveCodePane.GetSelection();
            ReorderParameters(selection);
        }

        private void OnRemoveParameterButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }
            var selection = IDE.ActiveCodePane.GetSelection();
            RemoveParameter(selection);
        }

        public void Rename(QualifiedSelection selection)
        {
            using (var view = new RenameDialog())
            {
                var parseResult = _parser.Parse(IDE.ActiveVBProject, this);
                var presenter = new RenamePresenter(IDE, view, parseResult, selection);
                presenter.Show();
            }
        }

        public void Rename(Declaration target)
        {
            using (var view = new RenameDialog())
            {
                var parseResult = _parser.Parse(IDE.ActiveVBProject, this);
                var presenter = new RenamePresenter(IDE, view, parseResult, new QualifiedSelection(target.QualifiedName.QualifiedModuleName, target.Selection));
                presenter.Show(target);
            }
        }

        public void ReorderParameters(QualifiedSelection selection)
        {
            using (var view = new ReorderParametersDialog())
            {
                var parseResult = _parser.Parse(IDE.ActiveVBProject);
                var presenter = new ReorderParametersPresenter(view, parseResult, selection);
                presenter.Show();
            }
        }

        public void RemoveParameter(QualifiedSelection selection)
        {
            var parseResult = _parser.Parse(IDE.ActiveVBProject);
            var presenter = new RemoveParameterPresenter(parseResult, selection);
        }

        private CommandBarButton AddMenuButton(CommandBarPopup menu)
        {
            return menu.Controls.Add(MsoControlType.msoControlButton) as CommandBarButton;
        }
    }
}
