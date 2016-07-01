using System;
using System.Diagnostics;
using System.Reflection;
using NLog;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.About
{
    public class AboutControlViewModel
    {
        public string Version
        {
            get
            {
                var name = Assembly.GetExecutingAssembly().GetName();
                return string.Format(RubberduckUI.Rubberduck_AboutBuild, name.Version, name.ProcessorArchitecture);
            }
        }

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
    }
}
