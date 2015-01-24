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
//todo: remove
using Rubberduck.Extensions;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class RubberduckMenu : IDisposable
    {
        private readonly VBE _vbe;

        private readonly TestMenu _testMenu; // todo: implement as DockablePresenter.
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly CodeExplorerMenu _codeExplorerMenu;
        private readonly CodeInspectionsMenu _codeInspectionsMenu;
        //private readonly RefactorMenu _refactorMenu; // todo: implement refactoring
        private readonly IConfigurationService _configService;

        public RubberduckMenu(VBE vbe, AddIn addIn, IConfigurationService configService, Parser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _configService = configService;

            _testMenu = new TestMenu(_vbe, addIn);
            _codeExplorerMenu = new CodeExplorerMenu(_vbe, addIn, parser);

            var todoSettings = configService.LoadConfiguration().UserSettings.ToDoListSettings;
            _todoItemsMenu = new ToDoItemsMenu(_vbe, addIn, todoSettings, parser);

            _codeInspectionsMenu = new CodeInspectionsMenu(_vbe, addIn, parser, inspections);
            //_refactorMenu = new RefactorMenu(_vbe, addIn);

        }

        public void Dispose()
        {
            _testMenu.Dispose();
            //_refactorMenu.Dispose();
        }

        private CommandBarButton _about;
        private CommandBarButton _settings;

        //todo: remove these temporary testing buttons
        private CommandBarButton _removeAll;
        private CommandBarButton _exportAll;
        private CommandBarButton _importAll;

        public void Initialize()
        {
            var menuBarControls = _vbe.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls);
            var menu = menuBarControls.Add(MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
            Debug.Assert(menu != null, "menu != null");

            menu.Caption = "Ru&bberduck";

            _testMenu.Initialize(menu.Controls);
            _codeExplorerMenu.Initialize(menu.Controls);
            //_refactorMenu.Initialize(menu.Controls);
            _todoItemsMenu.Initialize(menu.Controls);
            _codeInspectionsMenu.Initialize(menu.Controls);

            _settings = AddButton(menu, "&Options", true, new CommandBarButtonClickEvent(OnOptionsClick));
            _about = AddButton(menu, "&About...", true, new CommandBarButtonClickEvent(OnAboutClick));

            //todo: remove rapid prototyping buttons
            _removeAll = AddButton(menu, "RemoveAll", false, new CommandBarButtonClickEvent(OnRemoveClick));
            _exportAll = AddButton(menu, "ExportAll", false, new CommandBarButtonClickEvent(OnExportClick));
            _exportAll = AddButton(menu, "ImportAll", false, new CommandBarButtonClickEvent(OnImportClick));
            
        }

        //todo: remove
        private void OnImportClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var importDir = @"C:\Users\Christopher\Documents\SourceControlTest";
            var dirInfo = new System.IO.DirectoryInfo(importDir);

            foreach (var file in dirInfo.EnumerateFiles())
            {
                this._vbe.ActiveVBProject.VBComponents.ImportSourceFile(file.FullName);
            }
        }

        //todo: remove
        private void OnExportClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            var exportDir = @"C:\Users\Christopher\Documents\SourceControlTest";
            this._vbe.ActiveVBProject.ExportSourceFiles(exportDir);
        }

        //todo: remove
        private void OnRemoveClick(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            this._vbe.ActiveVBProject.RemoveAllComponents();
        }

        private CommandBarButton AddButton(CommandBarPopup parentMenu, string caption, bool beginGroup, CommandBarButtonClickEvent buttonClickHandler)
        {
            var button = parentMenu.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            button.Caption = caption;
            button.BeginGroup = beginGroup;
            button.Click += buttonClickHandler;

            return button;
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

        private int FindMenuInsertionIndex(CommandBarControls controls)
        {
            for (var i = 1; i <= controls.Count; i++)
            {
                // insert menu before "Window" built-in menu:
                if (controls[i].BuiltIn && controls[i].Caption == "&Window")
                {
                    return i;
                }
            }

            return controls.Count;
        }
    }
}