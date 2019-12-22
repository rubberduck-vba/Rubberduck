using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    public abstract class ComCommandBase : CommandBase
    {
        private readonly IVbeEvents _vbeEvents;
        
        protected ComCommandBase(IVbeEvents vbeEvents) 
        {
            _vbeEvents = vbeEvents;
            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute, true);
        }
        
        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return !_vbeEvents.Terminated;
        }
    }
}
