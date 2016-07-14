using NLog;
using Rubberduck.UI.RegexAssistant;
using System.Runtime.InteropServices;

namespace Rubberduck.UI.Command
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
