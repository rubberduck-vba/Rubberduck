using System.Runtime.InteropServices;
using Rubberduck.Settings;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all active code inspections for the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class InspectionResultsCommand : CommandBase
    {
        private readonly IPresenter _presenter;

        public InspectionResultsCommand(IPresenter presenter)
        {
            _presenter = presenter;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.InspectionResults; }
        }

        /// <summary>
        /// Runs code inspections 
        /// </summary>
        /// <param name="parameter"></param>
        public override void Execute(object parameter)
        {
            _presenter.Show();
        }
    }
}
