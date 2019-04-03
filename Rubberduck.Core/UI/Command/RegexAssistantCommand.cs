using System.Runtime.InteropServices;
using Rubberduck.UI.RegexAssistant;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the RegexAssistantDialog
    /// </summary>
    [ComVisible(false)]
    public class RegexAssistantCommand : CommandBase
    {
        protected override void OnExecute(object parameter)
        {
            using (var window = new RegexAssistantDialog())
            {
                window.ShowDialog();
            }
        }
    }
}
