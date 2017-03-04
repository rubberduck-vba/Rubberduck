using System;
using System.Diagnostics;
using System.Windows.Forms;
using NLog;
using Rubberduck.VersionCheck;

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
            : base(LogManager.GetCurrentClassLogger())
        {
            _versionCheck = versionCheck;
            _prompt = prompt;
            _process = process;
        }

        protected override async void ExecuteImpl(object parameter)
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
            if (_prompt.Show(prompt, RubberduckUI.Rubberduck, MessageBoxButtons.YesNo, MessageBoxIcon.Information) ==
                DialogResult.No)
            {
                return;
            }

            _process.Start("https://github.com/rubberduck-vba/Rubberduck/releases/latest");
        }
    }
}
