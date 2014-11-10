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
        private readonly RefactorMenu _refactorMenu;

        public App(VBE vbe)
        {
            _vbe = vbe;
            _testMenu = new TestMenu(_vbe);
            _refactorMenu = new RefactorMenu(_vbe);
        }

        public void Dispose()
        {
            _testMenu.Dispose();
            _refactorMenu.Dispose();
        }

        public void CreateExtUI()
        {
            _testMenu.Initialize();
            _refactorMenu.Initialize();
        }
    }
}
