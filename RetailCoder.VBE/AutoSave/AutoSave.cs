using System;
using System.IO;
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
                _timer.Enabled = _settings.IsEnabled;
            }
            if (e.PropertyName == "TimerDelay")
            {
                _timer.Interval = _settings.TimerDelay;
            }
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_vbe.VBProjects.OfType<VBProject>().Any(p => !p.Saved))
            {
                try
                {
                    // iterate to find if a file exists for each open project
                    // I do hope the compiler doesn't optimize this out
                    _vbe.VBProjects.OfType<VBProject>().Select(p => p.FileName).ToList();
                }
                catch (DirectoryNotFoundException)
                {
                    return;
                }

                _vbe.CommandBars.FindControl(Id: 3).Execute();
            }
        }

        public void Dispose()
        {
            _settings.PropertyChanged -= _settings_PropertyChanged;
            _timer.Elapsed -= _timer_Elapsed;

            _timer.Dispose();
        }
    }
}