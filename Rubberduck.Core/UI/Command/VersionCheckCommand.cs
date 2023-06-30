using System;
using System.Diagnostics;
using Rubberduck.Interaction;
using Rubberduck.VersionCheck;
using Rubberduck.Resources;
using Rubberduck.SettingsProvider;
using Rubberduck.Settings;
using System.Threading;

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
        private readonly IVersionCheckService _versionCheck;
        private readonly IMessageBox _prompt;
        private readonly IExternalProcess _process;
        IConfigurationService<Configuration> _config;

        public VersionCheckCommand(IVersionCheckService versionCheck, IMessageBox prompt, IExternalProcess process, IConfigurationService<Configuration> config)
        {
            _versionCheck = versionCheck;
            _prompt = prompt;
            _process = process;
            _config = config;
        }

        protected override async void OnExecute(object parameter)
        {
            var settings = _config.Read().UserSettings.GeneralSettings;
            if (_versionCheck.IsDebugBuild)
            {
                Logger.Info("Version check skipped for debug build.");
                return;
            }

            Logger.Info("Executing version check...");

            var tokenSource = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            await _versionCheck
                .GetLatestVersionAsync(settings, tokenSource.Token)
                .ContinueWith(t =>
                {
                    if (t.IsFaulted)
                    {
                        Logger.Warn(t.Exception);
                        return;
                    }

                    if (_versionCheck.CurrentVersion < t.Result)
                    {
                        var proceed = true;
                        if (_versionCheck.IsDebugBuild || !settings.IncludePreRelease)
                        {
                            // if the latest version has a revision number and isn't a pre-release build,
                            // avoid prompting since we can't know if the build already includes the latest version.
                            proceed = t.Result.Revision == 0;
                        }

                        if (proceed)
                        {
                            PromptAndBrowse(t.Result, settings.IncludePreRelease);
                        }
                        else
                        {
                            Logger.Info("Version check skips notification of an existing newer version available.");
                        }
                    }
                    else
                    {
                        Logger.Info("Version check completed: running current latest.");
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
