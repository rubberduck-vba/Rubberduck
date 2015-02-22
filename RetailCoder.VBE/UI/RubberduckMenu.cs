using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using CommandBarButtonClickEvent = Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Config;
using Rubberduck.Inspections;
using Rubberduck.UI.CodeInspections;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UI.CodeExplorer;
using Rubberduck.VBA;

namespace Rubberduck.UI
{
    public class RubberduckMenu : Menu
    {
        private readonly TestMenu _testMenu; // todo: implement as DockablePresenter.
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly CodeExplorerMenu _codeExplorerMenu;
        private readonly CodeInspectionsMenu _codeInspectionsMenu;
        private readonly RefactorMenu _refactorMenu;
        private readonly IConfigurationService _configService;

        public RubberduckMenu(VBE vbe, AddIn addIn, IConfigurationService configService, IRubberduckParser parser, IEnumerable<IInspection> inspections)
               :base(vbe, addIn)
        {
            _configService = configService;

            _testMenu = new TestMenu(vbe, addIn);

            var codeExplorer = new CodeExplorerWindow();
            var codePresenter = new CodeExplorerDockablePresenter(parser, vbe, addIn, codeExplorer);
            _codeExplorerMenu = new CodeExplorerMenu(vbe, addIn, codeExplorer, codePresenter);

            var todoSettings = configService.LoadConfiguration().UserSettings.ToDoListSettings;
            var todoExplorer = new ToDoExplorerWindow();
            var todoPresenter = new ToDoExplorerDockablePresenter(parser, todoSettings.ToDoMarkers, vbe, addIn, todoExplorer);
            _todoItemsMenu = new ToDoItemsMenu(vbe, addIn, todoExplorer, todoPresenter);

            var inspectionExplorer = new CodeInspections.CodeInspectionsWindow();
            var inspectionPresenter = new CodeInspectionsDockablePresenter(parser, inspections, vbe, addIn, inspectionExplorer);
            _codeInspectionsMenu = new CodeInspectionsMenu(vbe, addIn, inspectionExplorer, inspectionPresenter);

            _refactorMenu = new RefactorMenu(this.IDE, this.addInInstance, parser);
        }

        private CommandBarButton _about;
        private CommandBarButton _settings;
        private CommandBarButton _sourceControl;

        public void Initialize()
        {
            var menuBarControls = this.IDE.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls, "&Window");
            var menu = menuBarControls.Add(MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
            Debug.Assert(menu != null, "menu != null");

            menu.Caption = "Ru&bberduck";

            _testMenu.Initialize(menu.Controls);
            _codeExplorerMenu.Initialize(menu);
            _refactorMenu.Initialize(menu.Controls);
            _todoItemsMenu.Initialize(menu);
            _codeInspectionsMenu.Initialize(menu);

            //note: disabled for 1.2 release
            //_sourceControl = AddButton(menu, "Source Control", false, new CommandBarButtonClickEvent(OnSourceControlClick));

            _settings = AddButton(menu, "&Options", true, new CommandBarButtonClickEvent(OnOptionsClick));
            _about = AddButton(menu, "&About...", true, new CommandBarButtonClickEvent(OnAboutClick));

        }

        private void OnSourceControlClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new SourceControl.DummyGitView(this.IDE.ActiveVBProject))
            {
                window.ShowDialog();
            }
        }

        private void OnOptionsClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new Settings.SettingsDialog(_configService))
            {
                window.ShowDialog();
            }
        }

        private void OnAboutClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            using (var window = new AboutWindow())
            {
                window.ShowDialog();
            }
        }

        bool disposed = false;
        protected override void Dispose(bool disposing)
        {
            if (disposed)
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

            disposed = true;

            base.Dispose(disposing);
        }
    }
}