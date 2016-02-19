using System;
using System.IO;
using System.Linq;
using System.Timers;
using Microsoft.Vbe.Interop;
using Rubberduck.Settings;

namespace Rubberduck.AutoSave
{
    public class AutoSave : IDisposable
    {
        private readonly VBE _vbe;
        private readonly IGeneralConfigService _configService;
        private readonly Timer _timer = new Timer();
        private Configuration _config;

        private const int VbeSaveCommandId = 3;

        public AutoSave(VBE vbe, IGeneralConfigService configService)
        {
            _vbe = vbe;
            _configService = configService;
            _config = _configService.LoadConfiguration();

            _configService.SettingsChanged += ConfigServiceSettingsChanged;

            _timer.Enabled = _config.UserSettings.GeneralSettings.AutoSaveEnabled;
            _timer.Interval = _config.UserSettings.GeneralSettings.AutoSavePeriod * 1000;

            _timer.Elapsed += _timer_Elapsed;
        }

        void ConfigServiceSettingsChanged(object sender, EventArgs e)
        {
            _config = _configService.LoadConfiguration();

            _timer.Enabled = _config.UserSettings.GeneralSettings.AutoSaveEnabled;
            _timer.Interval = _config.UserSettings.GeneralSettings.AutoSavePeriod * 1000;
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
            _configService.LanguageChanged -= ConfigServiceSettingsChanged;
            _timer.Elapsed -= _timer_Elapsed;

            _timer.Dispose();
        }
    }
}