using System;
using System.Diagnostics;
using Rubberduck.Interaction;
using Rubberduck.VersionCheck;
using Rubberduck.Resources;
using Rubberduck.SettingsProvider;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    public interface IExternalProcess
    {
        void Start(string fileName);
    }

    public class ExternalProcess : IExternalProcess
    {
        public void Start(string fileName)
        {
            var info = new ProcessStartInfo(fileName)
            {
                WindowStyle = ProcessWindowStyle.Maximized
            };
            Process.Start(info);
        }
    }

    public class VersionCheckCommand : CommandBase
    {
        private readonly IVersionCheck _versionCheck;
        private readonly IMessageBox _prompt;
        private readonly IExternalProcess _process;
        IConfigurationService<Configuration> _config;

        public VersionCheckCommand(IVersionCheck versionCheck, IMessageBox prompt, IExternalProcess process, IConfigurationService<Configuration> config)
        {
            _versionCheck = versionCheck;
            _prompt = prompt;
            _process = process;
            _config = config;
        }

        protected override async void OnExecute(object parameter)
        {
            var settings = _config.Read().UserSettings.GeneralSettings;
            Logger.Info("Executing version check...");
            await _versionCheck
                .GetLatestVersionAsync(settings)
                .ContinueWith(t =>
                {
                    if (_versionCheck.CurrentVersion < t.Result)
                    {
                        PromptAndBrowse(t.Result, settings.IncludePreRelease);
                    }
                });
        }

        private void PromptAndBrowse(Version latestVersion, bool includePreRelease)
        {
            var buildType = includePreRelease 
                ? RubberduckUI.VersionCheck_BuildType_PreRelease 
                : RubberduckUI.VersionCheck_BuildType_Release;
            var prompt = string.Format(RubberduckUI.VersionCheck_NewVersionAvailable, _versionCheck.CurrentVersion, latestVersion, buildType);
            if (!_prompt.Question(prompt, RubberduckUI.Rubberduck))
            {
                return;
            }

            var url = includePreRelease
                ? "https://github.com/rubberduck-vba/Rubberduck/releases"
                : "https://github.com/rubberduck-vba/Rubberduck/releases/latest";
            _process.Start(url);
        }
    }
}
