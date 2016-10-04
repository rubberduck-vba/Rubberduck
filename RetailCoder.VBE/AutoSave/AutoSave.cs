using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Timers;
using Rubberduck.Settings;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

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
            _timer.Enabled = false;
        }

        public void ConfigServiceSettingsChanged(object sender, EventArgs e)
        {
            var config = _configService.LoadConfiguration();

            _timer.Enabled = config.UserSettings.GeneralSettings.AutoSaveEnabled
                && config.UserSettings.GeneralSettings.AutoSavePeriod != 0;

            _timer.Interval = config.UserSettings.GeneralSettings.AutoSavePeriod * 1000;
        }

        private void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            using (var projects = _vbe.VBProjects)
            if (projects.Any(p => !p.Saved))
            {
                try
                {
                    var foo = projects.Select(p => p.FileName).ToList();
                }
                catch (IOException)
                {
                    // note: VBProject.FileName getter throws IOException if unsaved
                    return;
                }

                var commandBars = _vbe.CommandBars;
                var control = commandBars.FindControl(Id: VbeSaveCommandId);
                control.Execute();
                Marshal.ReleaseComObject(control);
                Marshal.ReleaseComObject(commandBars);
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        private void Dispose(bool disposing)
        {
            if (!disposing)
            {
                return;
            }

            if (_configService != null)
            {
                _configService.SettingsChanged -= ConfigServiceSettingsChanged;
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
