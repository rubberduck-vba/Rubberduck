using System;
using System.Diagnostics;
using NLog;
using Rubberduck.UI.Command;
using Rubberduck.VersionCheck;

namespace Rubberduck.UI.About
{
    public class AboutControlViewModel
    {
        private readonly IVersionCheck _version;

        public AboutControlViewModel(IVersionCheck version)
        {
            _version = version;
        }

        public string Version => string.Format(RubberduckUI.Rubberduck_AboutBuild, _version.CurrentVersion);

        private CommandBase _uriCommand;
        public CommandBase UriCommand
        {
            get
            {
                if (_uriCommand != null)
                {
                    return _uriCommand;
                }
                return _uriCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), uri =>
                {
                    Process.Start(new ProcessStartInfo(((Uri)uri).AbsoluteUri));
                });
            }
        }

        public string AboutCopyright =>
            string.Format(RubberduckUI.AboutWindow_Copyright, DateTime.Now.Year);
    }
}
