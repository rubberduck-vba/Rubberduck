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
    [ComVisible(false)]
    public class RubberduckMenu : IDisposable
    {
        private readonly VBE _vbe;

        private readonly TestMenu _testMenu; // todo: implement as DockablePresenter.
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly CodeExplorerMenu _codeExplorerMenu;
        private readonly CodeInspectionsMenu _codeInspectionsMenu;
        //private readonly RefactorMenu _refactorMenu; // todo: implement refactoring

        public RubberduckMenu(VBE vbe, AddIn addIn, Configuration config, Parser parser, IEnumerable<IInspection> inspections)
        {
            _vbe = vbe;
            _testMenu = new TestMenu(_vbe, addIn);
            _codeExplorerMenu = new CodeExplorerMenu(_vbe, addIn, parser);
            _todoItemsMenu = new ToDoItemsMenu(_vbe, addIn, config.UserSettings.ToDoListSettings, parser);
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
            using (var window = new Settings.SettingsDialog())
            {
                window.ShowDialog();
            }
        }

        void OnAboutClick(CommandBarButton Ctrl, ref bool CancelDefault)
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