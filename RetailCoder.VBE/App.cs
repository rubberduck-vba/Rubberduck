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
    internal interface IApp
    {
        void CreateExtUI();
    }

    internal class App : IApp, IDisposable
    {
        private readonly VBE _vbe;
        private readonly TestMenu _testMenu;

        public App(VBE vbe)
        {
            _vbe = vbe;
            _testMenu = new TestMenu(_vbe);
        }

        public void Dispose()
        {
            _testMenu.Dispose();
        }

        public void CreateExtUI()
        {
            _testMenu.Initialize();
        }
    }
}
