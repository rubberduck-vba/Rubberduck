using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.UI.FindSymbol;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.UI.Refactorings.RemoveParameters;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI.Refactorings;

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
        private CommandBarButton _removeParametersButton;

        public void Initialize(CommandBarControls menuControls)
        {
            var menu = menuControls.Add(MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
            menu.Caption = RubberduckUI.RubberduckMenu_Refactor;

            _extractMethodButton = AddButton(menu, RubberduckUI.RefactorMenu_ExtractMethod, false, OnExtractMethodButtonClick);
            SetButtonImage(_extractMethodButton, Resources.ExtractMethod_6786_32, Resources.ExtractMethod_6786_32_Mask);

            _renameButton = AddButton(menu, RubberduckUI.RefactorMenu_Rename, false, OnRenameButtonClick);
            
            _reorderParametersButton = AddButton(menu, RubberduckUI.RefactorMenu_ReorderParameters, false, OnReorderParametersButtonClick, Resources.ReorderParameters_6780_32);
            SetButtonImage(_reorderParametersButton, Resources.ReorderParameters_6780_32, Resources.ReorderParameters_6780_32_Mask);

            _removeParametersButton = AddButton(menu, RubberduckUI.RefactorMenu_RemoveParameter, false, OnRemoveParameterButtonClick);
            SetButtonImage(_removeParametersButton, Resources.RemoveParameters_6781_32, Resources.RemoveParameters_6781_32_Mask);

            InitializeRefactorContextMenu();
        }

        private CommandBarPopup _refactorCodePaneContextMenu;
        public CommandBarPopup RefactorCodePaneContextMenu { get { return _refactorCodePaneContextMenu; } }

        private CommandBarButton _extractMethodContextButton;
        private CommandBarButton _renameContextButton;
        private CommandBarButton _reorderParametersContextButton;
        private CommandBarButton _removeParametersContextButton;

        private void InitializeRefactorContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _refactorCodePaneContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlPopup, Temporary: true, Before:beforeItem) as CommandBarPopup;
            _refactorCodePaneContextMenu.BeginGroup = true;
            _refactorCodePaneContextMenu.Caption = RubberduckUI.RubberduckMenu_Refactor;

            _extractMethodContextButton = AddButton(_refactorCodePaneContextMenu, RubberduckUI.RefactorMenu_ExtractMethod, false, OnExtractMethodButtonClick);
            SetButtonImage(_extractMethodContextButton, Resources.ExtractMethod_6786_32, Resources.ExtractMethod_6786_32_Mask);

            _renameContextButton = AddButton(_refactorCodePaneContextMenu, RubberduckUI.RefactorMenu_Rename, false, OnRenameButtonClick);

            _reorderParametersContextButton = AddButton(_refactorCodePaneContextMenu, RubberduckUI.RefactorMenu_ReorderParameters, false, OnReorderParametersButtonClick);
            SetButtonImage(_reorderParametersContextButton, Resources.ReorderParameters_6780_32, Resources.ReorderParameters_6780_32_Mask);

            _removeParametersContextButton = AddButton(_refactorCodePaneContextMenu, RubberduckUI.RefactorMenu_RemoveParameter, false, OnRemoveParameterButtonClick);
            SetButtonImage(_removeParametersContextButton, Resources.RemoveParameters_6781_32, Resources.RemoveParameters_6781_32_Mask);

            InitializeFindReferencesContextMenu(); //todo: untangle that mess...
            InitializeFindImplementationsContextMenu(); //todo: untangle that mess...
            InitializeFindSymbolContextMenu();
        }

        private CommandBarButton _findAllReferencesContextMenu;
        private void InitializeFindReferencesContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findAllReferencesContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            _findAllReferencesContextMenu.Caption = RubberduckUI.ContextMenu_FindAllReferences;
            _findAllReferencesContextMenu.Click += FindAllReferencesContextMenu_Click;
        }

        private CommandBarButton _findAllImplementationsContextMenu;
        private void InitializeFindImplementationsContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findAllImplementationsContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            _findAllImplementationsContextMenu.Caption = RubberduckUI.ContextMenu_GoToImplementation;
            _findAllImplementationsContextMenu.Click += FindAllImplementationsContextMenu_Click;
        }

        private CommandBarButton _findSymbolContextMenu;
        private void InitializeFindSymbolContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findSymbolContextMenu = IDE.CommandBars["Code Window"].Controls.Add(Type: MsoControlType.msoControlButton, Temporary: true, Before: beforeItem) as CommandBarButton;
            SetButtonImage(_findSymbolContextMenu, Resources.FindSymbol_6263_32, Resources.FindSymbol_6263_32_Mask);
            _findSymbolContextMenu.Caption = RubberduckUI.ContextMenu_FindSymbol;
            _findSymbolContextMenu.Click += FindSymbolContextMenuClick;
        }

        private void FindSymbolContextMenuClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FindSymbol();
        }

        private void FindSymbol()
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);
            var declarations = result.Declarations;
            var vm = new FindSymbolViewModel(declarations.Items.Where(item => !item.IsBuiltIn), _iconCache);
            using (var view = new FindSymbolDialog(vm))
            {
                view.Navigate += view_Navigate;
                view.ShowDialog();
            }
        }

        private void view_Navigate(object sender, NavigateCodeEventArgs e)
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

        private void FindAllReferencesContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FindAllReferences();
        }

        private void FindAllReferences()
        {
            var selection = IDE.ActiveCodePane.GetSelection();
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            var declarations = result.Declarations.Items.Where(item => item.DeclarationType != DeclarationType.ModuleOption
                && item.ComponentName == selection.QualifiedName.ComponentName)
                .ToList();

            var target = declarations.SingleOrDefault(item =>
                IsSelectedDeclaration(selection, item)
                || IsSelectedReference(selection, item));

            if (target != null)
            {
                FindAllReferences(target);
            }
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

        private void ShowImplementationsToolwindow(Declaration target, IEnumerable<Declaration> implementations, string name)
        {
            // throws a COMException if toolwindow was already closed
            var window = new SimpleListControl(string.Format(RubberduckUI.AllImplementations_Caption, name));
            var presenter = new ImplementationsListDockablePresenter(IDE, AddIn, window, implementations);
            presenter.Show();
        }

        private void ShowReferencesToolwindow(Declaration target)
        {
            // throws a COMException if toolwindow was already closed
            var window = new SimpleListControl(target);
            var presenter = new IdentifierReferencesListDockablePresenter(IDE, AddIn, window, target);
            presenter.Show();
        }

        private void FindAllImplementationsContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FindAllImplementations();
        }

        private void FindAllImplementations()
        {
            var selection = IDE.ActiveCodePane.GetSelection();
            var progress = new ParsingProgressPresenter();
            var parseResult = progress.Parse(_parser, IDE.ActiveVBProject);

            var implementsStatement = parseResult.Declarations.FindInterfaces()
                .SelectMany(i => i.References.Where(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
                .SingleOrDefault(r => r.QualifiedModuleName == selection.QualifiedName && r.Selection.Contains(selection.Selection));

            if (implementsStatement != null)
            {
                FindAllImplementations(implementsStatement.Declaration, parseResult);
            }

            var member = parseResult.Declarations.FindInterfaceImplementationMembers()
                    .SingleOrDefault(m => m.Project == selection.QualifiedName.Project
                                          && m.ComponentName == selection.QualifiedName.ComponentName
                                          && m.Selection.Contains(selection.Selection));

            if (member == null)
            {
                member = parseResult.Declarations.FindInterfaceMembers()
                    .SingleOrDefault(m => m.Project == selection.QualifiedName.Project
                                          && m.ComponentName == selection.QualifiedName.ComponentName
                                          && m.Selection.Contains(selection.Selection));
            }

            if (member == null)
            {
                return;
            }

            FindAllImplementations(member, parseResult);
        }

        public void FindAllImplementations(Declaration target)
        {
            var progress = new ParsingProgressPresenter();
            var parseResult = progress.Parse(_parser, IDE.ActiveVBProject);
            FindAllImplementations(target, parseResult);
        }

        public void FindAllImplementations(Declaration target, VBProjectParseResult parseResult)
        {
            IEnumerable<Declaration> implementations;
            string name;
            if (target.DeclarationType == DeclarationType.Class)
            {
                implementations = FindAllImplementationsOfClass(target, parseResult, out name);
            }
            else
            {
                implementations = FindAllImplementationsOfMember(target, parseResult, out name);
            }

            if (implementations == null)
            {
                implementations = new List<Declaration>();
            }

            var declarations = implementations as IList<Declaration> ?? implementations.ToList();
            var implementationsCount = declarations.Count();

            if (implementationsCount == 1)
            {
                // if there's only 1 implementation, just jump to it:
                ImplementationsListDockablePresenter.OnNavigateImplementation(IDE, declarations.First());
            }
            else if (implementationsCount > 1)
            {
                // if there's more than one implementation, show the dockable navigation window:
                try
                {
                    ShowImplementationsToolwindow(target, declarations, name);
                }
                catch (COMException)
                {
                    // the exception is related to the docked control host instance,
                    // trying again will work (I know, that's bad bad bad code)
                    ShowImplementationsToolwindow(target, declarations, name);
                }
            }
            else
            {
                var message = string.Format(RubberduckUI.AllImplementations_NoneFound, name);
                var caption = string.Format(RubberduckUI.AllImplementations_Caption, name);
                MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private IEnumerable<Declaration> FindAllImplementationsOfClass(Declaration target, VBProjectParseResult parseResult, out string name)
        {
            if (target.DeclarationType != DeclarationType.Class)
            {
                name = string.Empty;
                return null;
            }

            var result = target.References
                .Where(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext)
                .SelectMany(reference => parseResult.Declarations[reference.QualifiedModuleName.ComponentName])
                .ToList();

            name = target.ComponentName;
            return result;
        }

        private IEnumerable<Declaration> FindAllImplementationsOfMember(Declaration target, VBProjectParseResult parseResult, out string name)
        {
            if (!target.DeclarationType.HasFlag(DeclarationType.Member))
            {
                name = string.Empty;
                return null;
            }

            var isInterface = parseResult.Declarations.FindInterfaces()
                .Select(i => i.QualifiedName.QualifiedModuleName.ToString())
                .Contains(target.QualifiedName.QualifiedModuleName.ToString());

            if (isInterface)
            {
                name = target.ComponentName + "." + target.IdentifierName;
                return parseResult.Declarations.FindInterfaceImplementationMembers(target.IdentifierName)
                       .Where(item => item.IdentifierName == target.ComponentName + "_" + target.IdentifierName);
            }
            
            var member = parseResult.Declarations.FindInterfaceMember(target);
            name = member.ComponentName + "." + member.IdentifierName;
            return parseResult.Declarations.FindInterfaceImplementationMembers(member.IdentifierName)
                   .Where(item => item.IdentifierName == member.ComponentName + "_" + member.IdentifierName);
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            var isSameProject = declaration.Project == selection.QualifiedName.Project;
            var isSameModule = isSameProject && declaration.QualifiedName.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName;

            return declaration.References.Any(r =>
                isSameModule &&
                r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            var isSameProject = declaration.Project == selection.QualifiedName.Project;
            var isSameModule = isSameProject && declaration.QualifiedName.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName;

            // bug: QualifiedModuleName.Equals doesn't return expected value.
            return isSameModule && declaration.Selection.ContainsFirstCharacter(selection.Selection);
        }

        private void OnExtractMethodButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            ExtractMethod();
        }

        private void ExtractMethod()
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            var declarations = result.Declarations;
            var factory = new ExtractMethodPresenterFactory(_editor, declarations);
            var refactoring = new ExtractMethodRefactoring(factory, _editor);
            refactoring.InvalidSelection += refactoring_InvalidSelection;
            refactoring.Refactor();
        }

        void refactoring_InvalidSelection(object sender, EventArgs e)
        {
            MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void OnRenameButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }

            Rename();
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

        public void Rename()
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(IDE, view, result);
                var refactoring = new RenameRefactoring(factory);
                refactoring.Refactor();
            }
        }

        public void Rename(Declaration target)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(IDE, view, result);
                var refactoring = new RenameRefactoring(factory);
                refactoring.Refactor(target);
            }
        }

        public void ReorderParameters(QualifiedSelection selection)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new ReorderParametersDialog())
            {
                var presenter = new ReorderParametersPresenter(view, result, selection);
                presenter.Show();
            }
        }

        public void RemoveParameter(QualifiedSelection selection)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RemoveParametersDialog())
            {
                var presenter = new RemoveParametersPresenter(view, result, selection);
                presenter.Show();
            }
        }

        bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            _extractMethodButton.Click -= OnExtractMethodButtonClick;
            _extractMethodContextButton.Click -= OnExtractMethodButtonClick;
            _removeParametersButton.Click -= OnRemoveParameterButtonClick;
            _removeParametersContextButton.Click -= OnRemoveParameterButtonClick;
            _renameButton.Click -= OnRenameButtonClick;
            _renameContextButton.Click -= OnRenameButtonClick;
            _reorderParametersButton.Click -= OnReorderParametersButtonClick;
            _reorderParametersContextButton.Click -= OnReorderParametersButtonClick;
            _findAllReferencesContextMenu.Click -= FindAllReferencesContextMenu_Click;
            _findAllImplementationsContextMenu.Click -= FindAllImplementationsContextMenu_Click;
            _findSymbolContextMenu.Click -= FindSymbolContextMenuClick;

            _disposed = true;
            base.Dispose(disposing);
        }
    }
}
