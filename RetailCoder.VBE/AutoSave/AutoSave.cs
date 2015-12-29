using System;
using System.Threading;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.AutoSave
{
    public class AutoSave : IDisposable
    {
        private static IHostApplication _app;
        private static readonly Timer Timer = new Timer(Save);

        public AutoSave(IHostApplication app, uint time = 600000)
        {
            _app = app;
            Timer.Change(0, time);
        }

        public static void Save(object obj)
        {
            _app.Save();
        }

        public void Dispose()
        {
            Timer.Dispose();
        }
    }
}