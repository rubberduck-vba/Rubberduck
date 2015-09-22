using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Navigation;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.FindSymbol;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI
{
    public class RefactorMenu : Menu
    {
        private readonly IRubberduckParser _parser;
        private readonly IActiveCodePaneEditor _editor;
        private readonly IDeclarationNavigator _implementationsNavigator;
        private readonly IDeclarationNavigator _referenceNavigator;
        private readonly ICodePaneWrapperFactory _wrapperWrapperFactory;

        private readonly SearchResultIconCache _iconCache;

        public RefactorMenu(VBE vbe, AddIn addin, IRubberduckParser parser, IActiveCodePaneEditor editor, ICodePaneWrapperFactory wrapperFactory,
            IDeclarationNavigator implementationsNavigator,
            IDeclarationNavigator referenceNavigator)
            : base(vbe, addin)
        {
            _parser = parser;
            _editor = editor;
            _implementationsNavigator = implementationsNavigator;
            _referenceNavigator = referenceNavigator;
            _wrapperWrapperFactory = wrapperFactory;

            _iconCache = new SearchResultIconCache();
        }

        private CommandBarButton _extractMethodButton;
        private CommandBarButton _renameButton;
        private CommandBarButton _reorderParametersButton;
        private CommandBarButton _removeParametersButton;

        public void Initialize(CommandBarControls menuControls)
        {
            _menu = menuControls.Add(MsoControlType.msoControlPopup, Temporary: true) as CommandBarPopup;
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
                var codePane = _wrapperWrapperFactory.Create(e.QualifiedName.Component.CodeModule.CodePane);
                codePane.Selection = e.Selection;
            }
            catch (COMException)
            {
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void FindAllReferencesContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _referenceNavigator.Find();
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void FindAllImplementationsContextMenu_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            _implementationsNavigator.Find();
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnExtractMethodButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            new RefactorExtractMethodCommand(IDE, _parser, _editor).Execute(null);
        }

        void refactoring_InvalidSelection(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show(RubberduckUI.ExtractMethod_InvalidSelectionMessage, RubberduckUI.ExtractMethod_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
            var codePane = _wrapperWrapperFactory.Create(IDE.ActiveCodePane);
            var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);
            ReorderParameters(selection);
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnRemoveParameterButtonClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            new RefactorRemoveParametersCommand(IDE, _parser, _editor,_wrapperWrapperFactory).Execute(null)
        }

        public void Rename()
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(IDE, view, result, new MessageBox(), _wrapperWrapperFactory);
                var refactoring = new RenameRefactoring(factory, _editor, new MessageBox());
                refactoring.Refactor();
            }
        }

        public void Rename(Declaration target)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(IDE, view, result, new MessageBox(), _wrapperWrapperFactory);
                var refactoring = new RenameRefactoring(factory, _editor, new MessageBox());
                refactoring.Refactor(target);
            }
        }

        private void ReorderParameters(QualifiedSelection selection)
        {
            var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new ReorderParametersDialog())
            {
                var factory = new ReorderParametersPresenterFactory(_editor, view, result, new MessageBox());
                var refactoring = new ReorderParametersRefactoring(factory, _editor, new MessageBox());
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

            RemoveRefactorContextMenu();

            _disposed = true;
            base.Dispose(true);
        }
    }
}
