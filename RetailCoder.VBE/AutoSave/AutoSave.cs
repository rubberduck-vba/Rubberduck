using System;
using System.Linq;
using System.Timers;
using Microsoft.Vbe.Interop;

namespace Rubberduck.AutoSave
{
    public class AutoSave : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IAutoSaveSettings _settings;
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

        public AutoSave(VBE vbe, IAutoSaveSettings settings)
        {
            _vbe = vbe;
            _settings = settings;

            _settings.PropertyChanged += _settings_PropertyChanged;

            _timer.Enabled = _settings.IsEnabled;
            _timer.Interval = _settings.TimerDelay;

            _timer.Elapsed += _timer_Elapsed;
        }

        void _settings_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsEnabled")
            {
                IsEnabledChanged(sender, e);
            }
            if (e.PropertyName == "TimerDelay")
            {
                TimerDelayChanged(sender, e);
            }
        }

        void TimerDelayChanged(object sender, EventArgs e)
        {
            _timer.Interval = _settings.TimerDelay;
        }

        void IsEnabledChanged(object sender, EventArgs e)
        {
            _timer.Enabled = _settings.IsEnabled;
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
            _settings.IsEnabledChanged -= _settings_IsEnabledChanged;
            _settings.TimerDelayChanged -= _settings_TimerDelayChanged;

            _timer.Elapsed -= _timer_Elapsed;

            _timer.Dispose();
        }
    }
}