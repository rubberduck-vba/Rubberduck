using System;
using System.IO;
using System.Linq;
using System.Timers;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;

namespace Rubberduck.AutoSave
{
    public sealed class AutoSave : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IGeneralConfigService _configService;
        private Timer _timer = new Timer();

        private const int VbeSaveCommandId = 3;

        public AutoSave(VBE vbe, IGeneralConfigService configService)
        {
            _vbe = vbe;
            _configService = configService;

            _configService.SettingsChanged += ConfigServiceSettingsChanged;
            _timer.Elapsed += _timer_Elapsed;

            ConfigServiceSettingsChanged(null, EventArgs.Empty);
        }

        private void ConfigServiceSettingsChanged(object sender, EventArgs e)
        {
            var config = _configService.LoadConfiguration();

            _timer.Enabled = config.UserSettings.GeneralSettings.AutoSaveEnabled
                && config.UserSettings.GeneralSettings.AutoSavePeriod != 0;

            _timer.Interval = config.UserSettings.GeneralSettings.AutoSavePeriod * 1000;
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (_vbe.VBProjects.OfType<VBProject>().Any(p => !p.Saved))
            {
                try
                {
                    var projects = _vbe.VBProjects.OfType<VBProject>().Select(p => p.FileName).ToList();
                }
                catch (IOException)
                {
                    // note: VBProject.FileName getter throws IOException if unsaved
                    return;
                }

                _vbe.CommandBars.FindControl(Id: VbeSaveCommandId).Execute();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }

            if (_configService != null)
            {
                _configService.LanguageChanged -= ConfigServiceSettingsChanged;
            }

            if (_timer != null)
            {
                _timer.Elapsed -= _timer_Elapsed;
                _timer.Dispose();
                _timer = null;
            }
        }
    }
}
