using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.ToDoItems;
using Rubberduck.UnitTesting.UI;

namespace Rubberduck
{
    internal class RubberduckMenu : IDisposable
    {
        private readonly VBE _vbe;

        private readonly TestMenu _testMenu;
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly RefactorMenu _refactorMenu;

        public RubberduckMenu(VBE vbe, AddIn addIn, Config.Configuration config)
        {
            _vbe = vbe;
            _testMenu = new TestMenu(_vbe, addIn);
            _todoItemsMenu = new ToDoItemsMenu(_vbe, addIn, config.UserSettings.ToDoListSettings);
            _refactorMenu = new RefactorMenu(_vbe);
        }

        public void Dispose()
        {
            _testMenu.Dispose();
            _refactorMenu.Dispose();
        }

        private CommandBarButton _about;

        public void Initialize()
        {
            var menuBarControls = _vbe.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls);
            var menu = menuBarControls.Add(MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
            Debug.Assert(menu != null, "menu != null");

            menu.Caption = "Rubber&duck";

            _testMenu.Initialize(menu.Controls);
            _refactorMenu.Initialize(menu.Controls);
            _todoItemsMenu.Initialize(menu.Controls);

            _about = menu.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
            Debug.Assert(_about != null, "_about != null");

            _about.Caption = "&About...";
            _about.BeginGroup = true;
            _about.Click += OnAboutClick;
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