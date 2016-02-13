using System;
using System.ComponentModel;
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
        private readonly Timer _timer = new Timer();

        private const int VbeSaveCommandId = 3;

        public AutoSave(VBE vbe, IAutoSaveSettings settings)
        {
            _vbe = vbe;
            _settings = settings;

            _settings.PropertyChanged += _settings_PropertyChanged;

            _timer.Enabled = _settings.IsEnabled;
            _timer.Interval = _settings.TimerDelay;

            _timer.Elapsed += _timer_Elapsed;
        }

        private void _settings_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "IsEnabled":
                    _timer.Enabled = _settings.IsEnabled;
                    break;
                case "TimerDelay":
                    _timer.Interval = _settings.TimerDelay;
                    break;
            }
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_vbe.VBProjects.OfType<VBProject>().Any(p => !p.Saved))
            {
                try
                {
                    // note: VBProject.FileName getter throws IOException if unsaved
                    _vbe.VBProjects.OfType<VBProject>().Select(p => p.FileName).ToList();
                }
                catch (DirectoryNotFoundException)
                {
                    return;
                }

                _vbe.CommandBars.FindControl(Id: VbeSaveCommandId).Execute();
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