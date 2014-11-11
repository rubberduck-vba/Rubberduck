using Microsoft.Vbe.Interop;
using RetailCoderVBE.UnitTesting;
using RetailCoderVBE.UnitTesting.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE
{
    internal class App : IDisposable
    {
        private readonly VBE _vbe;
        private readonly TestMenu _testMenu;
        private AddIn _addInInst;
        private TaskList.TaskListMenu _taskListMenu;

        public App(VBE vbe, AddIn addInInst)
        {
            _addInInst = addInInst;
            _vbe = vbe;
            _testMenu = new TestMenu(_vbe, _addInInst);
            _taskListMenu = new TaskList.TaskListMenu(_vbe, _addInInst);
        }

        public void Dispose()
        {
            _testMenu.Dispose();
            _taskListMenu.Dispose();
        }

        public void CreateExtUI()
        {
            _testMenu.Initialize();
            _taskListMenu.Initialize();
        }
    }
}
