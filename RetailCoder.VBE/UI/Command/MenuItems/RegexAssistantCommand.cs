using System.Runtime.InteropServices;
using NLog;
using Rubberduck.UI.RegexAssistant;

namespace Rubberduck.UI.Command.MenuItems
{
    /// <summary>
    /// A command that displays the RegexAssistantDialog
    /// </summary>
    [ComVisible(false)]
    class RegexAssistantCommand : CommandBase
    {
        public RegexAssistantCommand() : base (LogManager.GetCurrentClassLogger())
        {
        }

        protected override void ExecuteImpl(object parameter)
        {
            using (var window = new RegexAssistantDialog())
            {
                window.ShowDialog();
            }
        }
    }
}
