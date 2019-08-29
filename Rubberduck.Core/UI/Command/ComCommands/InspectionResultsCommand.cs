using System.Runtime.InteropServices;
using Rubberduck.UI.Inspections;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    /// <summary>
    /// A command that runs all active code inspections for the active VBAProject.
    /// </summary>
    [ComVisible(false)]
    public class InspectionResultsCommand : ComCommandBase
    {
        private readonly InspectionResultsDockablePresenter _presenter;

        public InspectionResultsCommand(
            InspectionResultsDockablePresenter presenter, 
            IVbeEvents vbeEvents)
            : base(vbeEvents)
        {
            _presenter = presenter;
        }

        /// <summary>
        /// Runs code inspections 
        /// </summary>
        /// <param name="parameter"></param>
        protected override void OnExecute(object parameter)
        {
            _presenter.Show();
        }
    }
}
