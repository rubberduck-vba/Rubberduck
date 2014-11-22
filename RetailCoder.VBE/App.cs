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
            var config = Config.ConfigurationLoader.LoadConfiguration();
            _menu = new RubberduckMenu(vbe, addInInst, config);
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
