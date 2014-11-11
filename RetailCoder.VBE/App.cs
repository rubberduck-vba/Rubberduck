using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.ToDoItems;
using Rubberduck.UnitTesting;
using Rubberduck.UnitTesting.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck
{
    internal class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addInInst;
        private readonly RubberduckMenu _menu;

        public App(VBE vbe, AddIn addInInst)
        {
            _addInInst = addInInst;
            _vbe = vbe;
            _menu = new RubberduckMenu(_vbe, _addInInst);
        }

        public void Dispose()
        {
            _menu.Dispose();
        }

        public void CreateExtUI()
        {
            _menu.Initialize();
        }
    }

    internal class RubberduckMenu : IDisposable
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;

        private readonly TestMenu _testMenu;
        private readonly ToDoItemsMenu _todoItemsMenu;
        private readonly RefactorMenu _refactorMenu;

        public RubberduckMenu(VBE vbe, AddIn addIn)
        {
            _vbe = vbe;
            _addIn = addIn;
            _testMenu = new TestMenu(_vbe);
            _todoItemsMenu = new ToDoItemsMenu(_vbe, _addIn);
            _refactorMenu = new RefactorMenu(_vbe);
        }

        public void Dispose()
        {
            _testMenu.Dispose();
            _todoItemsMenu.Dispose();
            _refactorMenu.Dispose();
        }

        public void Initialize()
        {
            var menuBarControls = _vbe.CommandBars[1].Controls;
            var beforeIndex = FindMenuInsertionIndex(menuBarControls);
            var menu = menuBarControls.Add(Type: MsoControlType.msoControlPopup, Before: beforeIndex, Temporary: true) as CommandBarPopup;
            menu.Caption = "Rubber&duck";

            _testMenu.Initialize(menu.Controls);
            _refactorMenu.Initialize(menu.Controls);
            _todoItemsMenu.Initialize(menu.Controls);
        }

        private int FindMenuInsertionIndex(CommandBarControls controls)
        {
            for (int i = 1; i <= controls.Count; i++)
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
