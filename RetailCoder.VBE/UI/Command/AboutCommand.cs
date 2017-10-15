using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.About;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the About window.
    /// </summary>
    [ComVisible(false)]
    public class AboutCommand : CommandBase
    {
        public AboutCommand() : base(LogManager.GetCurrentClassLogger()) { }

        protected override void ExecuteImpl(object parameter)
        {
            using (var window = new AboutDialog())
            {
                window.ShowDialog();
            }
        }
    }
}
