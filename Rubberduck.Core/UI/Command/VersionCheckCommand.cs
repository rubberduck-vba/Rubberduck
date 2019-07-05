using System;
using System.Diagnostics;
using Rubberduck.Interaction;
using Rubberduck.VersionCheck;
using Rubberduck.Resources;

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

        public VersionCheckCommand(IVersionCheck versionCheck, IMessageBox prompt, IExternalProcess process)
        {
            _versionCheck = versionCheck;
            _prompt = prompt;
            _process = process;
        }

        protected override async void OnExecute(object parameter)
        {
            Logger.Info("Executing version check.");
            await _versionCheck
                .GetLatestVersionAsync()
                .ContinueWith(t =>
                {
                    if (_versionCheck.CurrentVersion < t.Result)
                    {
                        PromptAndBrowse(t.Result);
                    }
                });
        }

        private void PromptAndBrowse(Version latestVersion)
        {
            var prompt = string.Format(RubberduckUI.VersionCheck_NewVersionAvailable, latestVersion);
            if (!_prompt.Question(prompt, RubberduckUI.Rubberduck))
            {
                return;
            }

            _process.Start("https://github.com/rubberduck-vba/Rubberduck/releases/latest");
        }
    }
}
