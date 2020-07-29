using System.Runtime.InteropServices;
using Rubberduck.UI.About;
using Rubberduck.VersionCheck;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the About window.
    /// </summary>
    [ComVisible(false)]
    public class AboutCommand : CommandBase
    {
        public AboutCommand(IVersionCheck versionService, IWebNavigator web)
        {
            _versionService = versionService;
            _web = web;
        }

        private readonly IVersionCheck _versionService;
        private readonly IWebNavigator _web;

        protected override void OnExecute(object parameter)
        {
            using (var window = new AboutDialog(_versionService, _web))
            {
                window.ShowDialog();
            }
        }
    }
}
