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
        public AboutCommand(IVersionCheck versionService)
        {
            _versionService = versionService;
        }

        private readonly IVersionCheck _versionService;

        protected override void OnExecute(object parameter)
        {
            using (var window = new AboutDialog(_versionService))
            {
                window.ShowDialog();
            }
        }
    }
}
