using System.Runtime.InteropServices;
using Rubberduck.UI.Command.ComCommands;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]    
    public class RefreshCommand : CommandBase
    {
        private readonly ReparseCommand _command;

        public RefreshCommand(ReparseCommand command) 
        {
            _command = command;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _command?.CanExecute(parameter) ?? false;
        }

        protected override void OnExecute(object parameter)
        {
            _command?.Execute(parameter);
        }
    }
}
