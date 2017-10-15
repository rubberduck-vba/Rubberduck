using System;
using System.Linq;
using System.Timers;
using Rubberduck.Settings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoSave
{
    public sealed class AutoSave : IDisposable
    {
        private readonly IVBE _vbe;
        private readonly IGeneralConfigService _configService;
        private Timer _timer = new Timer();

        private const int VbeSaveCommandId = 3;

        public AutoSave(IVBE vbe, IGeneralConfigService configService)
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
            SaveAllUnsavedProjects();
        }

            private void SaveAllUnsavedProjects()
            {
                var saveCommand = _vbe.CommandBars.FindControl(VbeSaveCommandId);
                var activeProject = _vbe.ActiveVBProject;
                var unsaved = _vbe
                    .VBProjects
                    .Where(project => !project.IsSaved && !string.IsNullOrEmpty(project.FileName));

                foreach (var project in unsaved)
                {
                    _vbe.ActiveVBProject = project;
                    saveCommand.Execute();
                }

                _vbe.ActiveVBProject = activeProject;
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
