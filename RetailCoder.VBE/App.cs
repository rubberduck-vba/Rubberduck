using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;

namespace Rubberduck
{
    [ComVisible(false)]
    public class App : IDisposable
    {
        private readonly RubberduckMenu _menu;

        public App(VBE vbe, AddIn addInInst)
        {
            _menu = new RubberduckMenu(vbe, addInInst);
        }

        public void Dispose()
        {
            _menu.Dispose();
        }

        public void CreateExtUi()
        {
            _menu.Initialize();
        }
    }
}
