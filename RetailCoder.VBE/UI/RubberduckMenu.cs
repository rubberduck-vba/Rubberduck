using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.ToDoItems;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class RubberduckMenu : IDisposable
    {
        private readonly VBE _vbe;

        private readonly TestMenu _testMenu; // todo: implement as DockablePresenter.
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly RefactorMenu _refactorMenu;

        public RubberduckMenu(VBE vbe, AddIn addIn, Config.Configuration config)
        {
            _vbe = vbe;
            _testMenu = new TestMenu(_vbe, addIn);
            _todoItemsMenu = new ToDoItemsMenu(_vbe, addIn, config.UserSettings.ToDoListSettings, new Parser());
            _refactorMenu = new RefactorMenu(_vbe, addIn);
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

            menu.Caption = "Ru&bberduck";

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