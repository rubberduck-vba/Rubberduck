using System;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Ninject;
using Ninject.Parameters;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Settings;
using Rubberduck.ToDoItems;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.Settings;
using Rubberduck.UI.SourceControl;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.UI
{
    internal class RubberduckMenu : Menu
    {
        private readonly TestMenu _testMenu;
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly CodeExplorerMenu _codeExplorerMenu;
        private readonly CodeInspectionsMenu _codeInspectionsMenu;
        private readonly RefactorMenu _refactorMenu;
        private readonly IGeneralConfigService _configService;
        private readonly IRubberduckParser _parser;
        private readonly IActiveCodePaneEditor _editor;
        private readonly IRubberduckCodePaneFactory _factory;
        private readonly AddIn _addIn;

        private CommandBarButton _about;
        private CommandBarButton _settings;
        private CommandBarButton _sourceControl;

        private ProjectExplorerContextMenu _projectExplorerContextMenu;

        public RubberduckMenu(VBE vbe, AddIn addIn, IGeneralConfigService configService, IRubberduckParser parser, IActiveCodePaneEditor editor, IInspector inspector, IRubberduckCodePaneFactory factory)
            : base(vbe, addIn)
        {
            _addIn = addIn;
            _parser = parser;
            _editor = editor;
            _factory = factory;
            _configService = configService;

            var testExplorer = new TestExplorerWindow();
            var testEngine = new TestEngine();
            var testGridViewSort = new GridViewSort<TestExplorerItem>(RubberduckUI.Result, false);
            var testPresenter = new TestExplorerDockablePresenter(vbe, addIn, testExplorer, testEngine, testGridViewSort, _factory);
            _testMenu = new TestMenu(vbe, addIn, testExplorer, testPresenter);

            var codeExplorer = new CodeExplorerWindow();
            var codePresenter = new CodeExplorerDockablePresenter(parser, vbe, addIn, codeExplorer, _factory);
            codePresenter.RunAllTests += CodePresenterRunAllAllTests;
            codePresenter.RunInspections += codePresenter_RunInspections;
            codePresenter.Rename += codePresenter_Rename;
            codePresenter.FindAllReferences += codePresenter_FindAllReferences;
            codePresenter.FindAllImplementations += codePresenter_FindAllImplementations;
            _codeExplorerMenu = new CodeExplorerMenu(vbe, addIn, codeExplorer, codePresenter);

            var todoSettings = configService.LoadConfiguration().UserSettings.ToDoListSettings;
            var todoExplorer = new ToDoExplorerWindow();
            var todoGridViewSort = new GridViewSort<ToDoItem>(RubberduckUI.Priority, false);
            var todoPresenter = new ToDoExplorerDockablePresenter(parser, todoSettings.ToDoMarkers, vbe, addIn, todoExplorer, todoGridViewSort, _factory);
            _todoItemsMenu = new ToDoItemsMenu(vbe, addIn, todoExplorer, todoPresenter);

            var inspectionExplorer = new CodeInspectionsWindow();
            var inspectionGridViewSort = new GridViewSort<CodeInspectionResultGridViewItem>(RubberduckUI.Component, false);
            var inspectionPresenter = new CodeInspectionsDockablePresenter(inspector, vbe, addIn, inspectionExplorer, inspectionGridViewSort, _factory);
            _codeInspectionsMenu = new CodeInspectionsMenu(vbe, addIn, inspectionExplorer, inspectionPresenter);

            _refactorMenu = new RefactorMenu(IDE, AddIn, parser, editor, _factory);
        }

        private void codePresenter_FindAllReferences(object sender, NavigateCodeEventArgs e)
        {
            _refactorMenu.FindAllReferences(e.Declaration);
        }

        private void codePresenter_FindAllImplementations(object sender, NavigateCodeEventArgs e)
        {
            _refactorMenu.FindAllImplementations(e.Declaration);
        }

        private void codePresenter_Rename(object sender, TreeNodeNavigateCodeEventArgs e)
        {
            var declaration = e.Node.Tag as Declaration;
            if (declaration == null)
            {
                return;
            }

            _refactorMenu.Rename(declaration);
        }

        private void codePresenter_RunInspections(object sender, EventArgs e)
        {
            _codeInspectionsMenu.Inspect();
        }

        private void CodePresenterRunAllAllTests(object sender, EventArgs e)
        {
            _testMenu.RunAllTests();
        }

        public void Initialize()
        {
            const int windowMenuId = 30009;
            var menuBarControls = IDE.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls, windowMenuId);
            _menu = menuBarControls.Add(MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;

            _menu.Caption = RubberduckUI.RubberduckMenu;

            _testMenu.Initialize(_menu.Controls);
            _codeExplorerMenu.Initialize(_menu);
            _refactorMenu.Initialize(_menu.Controls);
            _todoItemsMenu.Initialize(_menu);
            _codeInspectionsMenu.Initialize(_menu);

            _sourceControl = AddButton(_menu, RubberduckUI.RubberduckMenu_SourceControl, false, OnSourceControlClick);
            _settings = AddButton(_menu, RubberduckUI.RubberduckMenu_Options, true, OnOptionsClick);
            _about = AddButton(_menu, RubberduckUI.RubberduckMenu_About, true, OnAboutClick);

            _projectExplorerContextMenu = new ProjectExplorerContextMenu(IDE, _addIn, _parser, _editor, _factory);
            _projectExplorerContextMenu.Initialize();
            _projectExplorerContextMenu.RunInspections += codePresenter_RunInspections;
            _projectExplorerContextMenu.FindReferences += codePresenter_FindAllReferences;
            _projectExplorerContextMenu.FindImplementations += codePresenter_FindAllImplementations;
            _projectExplorerContextMenu.RunAllTests += CodePresenterRunAllAllTests;
        }

        private SourceControlPresenter _sourceControlPresenter;
        //I'm not the one with the bad name, MS is. Signature must match delegate definition.
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnSourceControlClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_sourceControlPresenter == null)
            {
                var kernel = new StandardKernel(new SourceControlBindings());
                var vbeArg = new ConstructorArgument("vbe", this.IDE);
                var addinArg = new ConstructorArgument("addin", this.AddIn);

                _sourceControlPresenter = kernel.Get<SourceControlPresenter>(vbeArg, addinArg);
            }

            _sourceControlPresenter.Show();
        }

        //I'm not the one with the bad name, MS is. Signature must match delegate definition.
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnOptionsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new _SettingsDialog(_configService))
            {
                window.ShowDialog();
            }
        }

        //I'm not the one with the bad name, MS is. Signature must match delegate definition.
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private void OnAboutClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new _AboutWindow())
            {
                window.ShowDialog();
            }
        }

        private bool _disposed;
        private CommandBarPopup _menu;

        protected override void Dispose(bool disposing)
        {
            if (_disposed || !disposing)
            {
                return;
            }

            if (_projectExplorerContextMenu != null)
            {
                _projectExplorerContextMenu.Dispose();
            }

            if (_todoItemsMenu != null)
            {
                _todoItemsMenu.Dispose();
            }

            if (_refactorMenu != null)
            {
                _refactorMenu.Dispose();
            }

            if (_codeExplorerMenu != null)
            {
                _codeExplorerMenu.Dispose();
            }

            if (_testMenu != null)
            {
                _testMenu.Dispose();
            }

            if (_codeInspectionsMenu != null)
            {
                _codeInspectionsMenu.Dispose();
            }

            _about.Click -= OnAboutClick;
            _settings.Click -= OnOptionsClick;
            _sourceControl.Click -= OnSourceControlClick;

            var menuBarControls = IDE.CommandBars[1].Controls;
            menuBarControls.Parent.FindControl(_menu.Type, _menu.Id, _menu.Tag, _menu.Visible).Delete();

            _disposed = true;
            base.Dispose(true);
        }
    }
}