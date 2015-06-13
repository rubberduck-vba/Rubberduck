﻿using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.VBIDEApi;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.FindSymbol;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

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
            _menu = menuControls.Add(MsoControlType.msoControlPopup, null, null, null, true) as CommandBarPopup;
            _menu.Caption = RubberduckUI.RubberduckMenu_Refactor;

            _extractMethodButton = AddButton(_menu, RubberduckUI.RefactorMenu_ExtractMethod, false, OnExtractMethodButtonClick);
            SetButtonImage(_extractMethodButton, Resources.ExtractMethod_6786_32, Resources.ExtractMethod_6786_32_Mask);

            _renameButton = AddButton(_menu, RubberduckUI.RefactorMenu_Rename, false, OnRenameButtonClick);
            
            _reorderParametersButton = AddButton(_menu, RubberduckUI.RefactorMenu_ReorderParameters, false, OnReorderParametersButtonClick, Resources.ReorderParameters_6780_32);
            SetButtonImage(_reorderParametersButton, Resources.ReorderParameters_6780_32, Resources.ReorderParameters_6780_32_Mask);

            _removeParametersButton = AddButton(_menu, RubberduckUI.RefactorMenu_RemoveParameter, false, OnRemoveParameterButtonClick);
            SetButtonImage(_removeParametersButton, Resources.RemoveParameters_6781_32, Resources.RemoveParameters_6781_32_Mask);

            InitializeRefactorContextMenu();
        }

        private CommandBarPopup _refactorCodePaneContextMenu;

        private CommandBarButton _extractMethodContextButton;
        private CommandBarButton _renameContextButton;
        private CommandBarButton _reorderParametersContextButton;
        private CommandBarButton _removeParametersContextButton;

        private void InitializeRefactorContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _refactorCodePaneContextMenu = IDE.CommandBars["Code Window"].Controls.Add(MsoControlType.msoControlPopup, null, null, beforeItem, true) as CommandBarPopup;
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

        private void RemoveRefactorContextMenu()
        {
            _extractMethodContextButton.Delete();
            _renameContextButton.Delete();
            _reorderParametersContextButton.Delete();
            _removeParametersContextButton.Delete();
            _refactorCodePaneContextMenu.Delete();

            _findAllReferencesContextMenu.Delete();
            _findAllImplementationsContextMenu.Delete();
            _findSymbolContextMenu.Delete();
        }

        private CommandBarButton _findAllReferencesContextMenu;
        private void InitializeFindReferencesContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findAllReferencesContextMenu = IDE.CommandBars["Code Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem, true) as CommandBarButton;
            _findAllReferencesContextMenu.Caption = RubberduckUI.ContextMenu_FindAllReferences;
            _findAllReferencesContextMenu.ClickEvent += FindAllReferencesContextMenu_Click;
        }

        private CommandBarButton _findAllImplementationsContextMenu;
        private void InitializeFindImplementationsContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findAllImplementationsContextMenu = IDE.CommandBars["Code Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem, true) as CommandBarButton;
            _findAllImplementationsContextMenu.Caption = RubberduckUI.ContextMenu_GoToImplementation;
            _findAllImplementationsContextMenu.ClickEvent += FindAllImplementationsContextMenu_Click;
        }

        private CommandBarButton _findSymbolContextMenu;
        private void InitializeFindSymbolContextMenu()
        {
            var beforeItem = IDE.CommandBars["Code Window"].Controls.Cast<CommandBarControl>().First(control => control.Id == 2529).Index;
            _findSymbolContextMenu = IDE.CommandBars["Code Window"].Controls.Add(MsoControlType.msoControlButton, null, null, beforeItem, true) as CommandBarButton;
            SetButtonImage(_findSymbolContextMenu, Resources.FindSymbol_6263_32, Resources.FindSymbol_6263_32_Mask);
            _findSymbolContextMenu.Caption = RubberduckUI.ContextMenu_FindSymbol;
            _findSymbolContextMenu.ClickEvent += FindSymbolContextMenuClick;
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
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

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void FindAllReferencesContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FindAllReferences();
        }

        public void FindAllReferences()
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

        private void ShowImplementationsToolwindow(IEnumerable<Declaration> implementations, string name)
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

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void FindAllImplementationsContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            FindAllImplementations();
        }

        public void FindAllImplementations()
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
                                      && m.Selection.Contains(selection.Selection)) ??
                         parseResult.Declarations.FindInterfaceMembers()
                                          .SingleOrDefault(m => m.Project == selection.QualifiedName.Project
                                                                && m.ComponentName == selection.QualifiedName.ComponentName
                                                                && m.Selection.Contains(selection.Selection));

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

        private void FindAllImplementations(Declaration target, VBProjectParseResult parseResult)
        {
            string name;
            var implementations = (target.DeclarationType == DeclarationType.Class
                ? FindAllImplementationsOfClass(target, parseResult, out name)
                : FindAllImplementationsOfMember(target, parseResult, out name)) ??
                                  new List<Declaration>();

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
                    ShowImplementationsToolwindow(declarations, name);
                }
                catch (COMException)
                {
                    // the exception is related to the docked control host instance,
                    // trying again will work (I know, that's bad bad bad code)
                    ShowImplementationsToolwindow(declarations, name);
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

        [SuppressMessage("ReSharper", "InconsistentNaming")]
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

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRenameButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }

            Rename();
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnReorderParametersButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (IDE.ActiveCodePane == null)
            {
                return;
            }
            var selection = IDE.ActiveCodePane.GetSelection();
            ReorderParameters(selection);
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
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

        private void ReorderParameters(QualifiedSelection selection)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new ReorderParametersDialog())
            {
                var factory = new ReorderParametersPresenterFactory(_editor, view, result);
                var refactoring = new ReorderParametersRefactoring(factory);
                refactoring.Refactor(selection);
            }
        }

        private void RemoveParameter(QualifiedSelection selection)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RemoveParametersDialog())
            {
                var factory = new RemoveParametersPresenterFactory(_editor, view, result);
                var refactoring = new RemoveParametersRefactoring(factory);
                refactoring.Refactor(selection);
            }
        }

        bool _disposed;
        private CommandBarPopup _menu;

        protected override void Dispose(bool disposing)
        {
            if (_disposed || !disposing)
            {
                return;
            }

            _extractMethodButton.ClickEvent -= OnExtractMethodButtonClick;
            _extractMethodContextButton.ClickEvent -= OnExtractMethodButtonClick;
            _removeParametersButton.ClickEvent -= OnRemoveParameterButtonClick;
            _removeParametersContextButton.ClickEvent -= OnRemoveParameterButtonClick;
            _renameButton.ClickEvent -= OnRenameButtonClick;
            _renameContextButton.ClickEvent -= OnRenameButtonClick;
            _reorderParametersButton.ClickEvent -= OnReorderParametersButtonClick;
            _reorderParametersContextButton.ClickEvent -= OnReorderParametersButtonClick;
            _findAllReferencesContextMenu.ClickEvent -= FindAllReferencesContextMenu_Click;
            _findAllImplementationsContextMenu.ClickEvent -= FindAllImplementationsContextMenu_Click;
            _findSymbolContextMenu.ClickEvent -= FindSymbolContextMenuClick;

            RemoveRefactorContextMenu();

            _disposed = true;
            base.Dispose(true);
        }
    }
}
