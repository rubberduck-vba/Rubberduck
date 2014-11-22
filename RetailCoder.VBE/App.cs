using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;

namespace Rubberduck
{
    [ComVisible(false)]
    public class App : IDisposable
    {
        private readonly RubberduckMenu _menu;
        private Config.Configuration _config;

        public App(VBE vbe, AddIn addInInst)
        {
            _config = Config.ConfigurationLoader.LoadConfiguration();
            _menu = new RubberduckMenu(vbe, addInInst, _config);
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
