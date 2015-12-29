using System;
using System.Linq;
using System.Timers;
using Microsoft.Vbe.Interop;

namespace Rubberduck.AutoSave
{
    public interface IAutoSave
    {
        bool IsEnabled { get; set; }
        double TimerDelay { get; set; }
    }

    public class AutoSave : IAutoSave, IDisposable
    {
        private readonly VBE _vbe;
        // ReSharper disable once InconsistentNaming
        private readonly Timer _timer = new Timer();

        public bool IsEnabled
        {
            get { return _timer.Enabled; }
            set { _timer.Enabled = value; }
        }

        public double TimerDelay
        {
            get { return _timer.Interval; }
            set { _timer.Interval = value; }
        }

        public AutoSave(VBE vbe, uint time = 600000)
        {
            _vbe = vbe;
            _timer.Interval = time;
            _timer.Elapsed += _timer_Elapsed;
            _timer.Start();
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_vbe.VBProjects.OfType<VBProject>().Any(p => !p.Saved))
            {
                _vbe.CommandBars.FindControl(Id: 3).Execute();
            }
        }

        public void Dispose()
        {
            _timer.Dispose();
        }
    }
}