using Microsoft.Vbe.Interop;
using System;

namespace Rubberduck
{
    internal class App : IDisposable
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
