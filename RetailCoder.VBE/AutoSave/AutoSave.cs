using System;
using System.Threading;
using Microsoft.Vbe.Interop;

namespace Rubberduck.AutoSave
{
    public class AutoSave : IDisposable
    {
        private static VBE _vbe;
        private static readonly Timer Timer = new Timer(Save);

        public AutoSave(VBE vbe, uint time)
        {
            _vbe = vbe;
            Timer.Change(0, time);
        }

        public static void Save(object foo)
        {
            _vbe.ActiveVBProject.SaveAs(_vbe.ActiveVBProject.Name + "_" + DateTime.Now);
        }

        public void Dispose()
        {
            Timer.Dispose();
        }
    }
}