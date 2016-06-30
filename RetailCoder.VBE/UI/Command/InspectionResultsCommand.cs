using System.Runtime.InteropServices;
using NLog;
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

        public InspectionResultsCommand(IPresenter presenter) : base(LogManager.GetCurrentClassLogger())
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
        protected override void ExecuteImpl(object parameter)
        {
            _presenter.Show();
        }
    }
}
