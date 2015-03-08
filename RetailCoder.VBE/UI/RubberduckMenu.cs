using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Inspections;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.Settings;
using Rubberduck.UI.SourceControl;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;
using Rubberduck.VBA;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;

namespace Rubberduck.UI
{
    public class RubberduckMenu : Menu
    {
        private readonly TestMenu _testMenu; 
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly CodeExplorerMenu _codeExplorerMenu;
        private readonly CodeInspectionsMenu _codeInspectionsMenu;
        private readonly RefactorMenu _refactorMenu;
        private readonly IConfigurationService _configService;

        public RubberduckMenu(VBE vbe, AddIn addIn, IConfigurationService configService, IRubberduckParser parser, IInspector inspector)
            : base(vbe, addIn)
        {
            _configService = configService;

            var testExplorer = new TestExplorerWindow();
            var testEngine = new TestEngine();
            var testPresenter = new TestExplorerDockablePresenter(vbe, addIn, testExplorer, testEngine);
            _testMenu = new TestMenu(vbe, addIn, testExplorer, testPresenter);

            var codeExplorer = new CodeExplorerWindow();
            var codePresenter = new CodeExplorerDockablePresenter(parser, vbe, addIn, codeExplorer);
            codePresenter.RunAllTests += codePresenter_RunAllTests;
            codePresenter.RunInspections += codePresenter_RunInspections;
            _codeExplorerMenu = new CodeExplorerMenu(vbe, addIn, codeExplorer, codePresenter);

            var todoSettings = configService.LoadConfiguration().UserSettings.ToDoListSettings;
            var todoExplorer = new ToDoExplorerWindow();
            var todoPresenter = new ToDoExplorerDockablePresenter(parser, todoSettings.ToDoMarkers, vbe, addIn, todoExplorer);
            _todoItemsMenu = new ToDoItemsMenu(vbe, addIn, todoExplorer, todoPresenter);

            var inspectionExplorer = new CodeInspectionsWindow();
            var inspectionPresenter = new CodeInspectionsDockablePresenter(inspector, vbe, addIn, inspectionExplorer);
            _codeInspectionsMenu = new CodeInspectionsMenu(vbe, addIn, inspectionExplorer, inspectionPresenter);

            _refactorMenu = new RefactorMenu(IDE, AddIn, parser);
        }

        private void codePresenter_RunInspections(object sender, System.EventArgs e)
        {
            _codeInspectionsMenu.Inspect();
        }

        private void codePresenter_RunAllTests(object sender, System.EventArgs e)
        {
            _testMenu.RunAllTests();
        }

        public void Initialize()
        {
            const int windowMenuId = 30009;
            var menuBarControls = IDE.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls, windowMenuId);
            var menu = menuBarControls.Add(MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
            Debug.Assert(menu != null, "menu != null");

            menu.Caption = "Ru&bberduck";

            _testMenu.Initialize(menu.Controls);
            _codeExplorerMenu.Initialize(menu);
            _refactorMenu.Initialize(menu.Controls);
            _todoItemsMenu.Initialize(menu);
            _codeInspectionsMenu.Initialize(menu);

            AddButton(menu, "Source Control", false, OnSourceControlClick);
            AddButton(menu, "&Options", true, OnOptionsClick);
            AddButton(menu, "&About...", true, OnAboutClick);
        }

        private void OnSourceControlClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new DummyGitView(IDE.ActiveVBProject))
            {
                window.ShowDialog();
            }
        }

        private void OnOptionsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new _SettingsDialog(_configService))
            {
                window.ShowDialog();
            }
        }

        private void OnAboutClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new _AboutWindow())
            {
                window.ShowDialog();
            }
        }

        bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
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
            }

            _disposed = true;

            base.Dispose(disposing);
        }
    }
}