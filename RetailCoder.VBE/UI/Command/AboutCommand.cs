using System.Runtime.InteropServices;
using Rubberduck.UI.About;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the About window.
    /// </summary>
    [ComVisible(false)]
    public class AboutCommand : CommandBase
    {
        public override void Execute(object parameter)
        {
            using (var window = new AboutDialog())
            {
                window.ShowDialog();
            }
        }
    }
}
