using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Input;
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

        private ICommand _uriCommand;
        public ICommand UriCommand
        {
            get
            {
                if (_uriCommand != null)
                {
                    return _uriCommand;
                }
                return _uriCommand = new DelegateCommand(uri =>
                {
                    Process.Start(new ProcessStartInfo(((Uri)uri).AbsoluteUri));
                });
            }
        }
    }
}
