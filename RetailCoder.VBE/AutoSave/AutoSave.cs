using System;
using System.Linq;
using System.Threading;
using Microsoft.Vbe.Interop;

namespace Rubberduck.AutoSave
{
    public class AutoSave : IDisposable
    {
        private static VBE _vbe;
        private static readonly Timer Timer = new Timer(Save);

        public AutoSave(VBE vbe, uint time = 600000)
        {
            _vbe = vbe;
            Timer.Change(0, time);
        }

        public static void Save(object obj)
        {
            if (_vbe.VBProjects.OfType<VBProject>().Any(p => !p.Saved))
            {
                _vbe.CommandBars.FindControl(Id: 3).Execute();
            }
        }

        public void Dispose()
        {
            Timer.Dispose();
        }
    }
}