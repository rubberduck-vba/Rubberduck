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

            _configService.SettingsChanged += ConfigServiceSettingsChanged;

            // todo: move this out of ctor
            //_timer.Enabled = _config.UserSettings.GeneralSettings.AutoSaveEnabled 
            //    && _config.UserSettings.GeneralSettings.AutoSavePeriod != 0;

            //if (_config.UserSettings.GeneralSettings.AutoSavePeriod != 0)
            //{
            //    _timer.Interval = _config.UserSettings.GeneralSettings.AutoSavePeriod * 1000;
            //    _timer.Elapsed += _timer_Elapsed;
            //}
        }

        private void ConfigServiceSettingsChanged(object sender, EventArgs e)
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
            _configService.LanguageChanged -= ConfigServiceSettingsChanged;
            _timer.Elapsed -= _timer_Elapsed;

            _timer.Dispose();
        }
    }
}