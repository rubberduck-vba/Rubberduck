using System.Runtime.InteropServices;
using NLog;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]    
    public class RefreshCommand : CommandBase
    {
        private readonly ReparseCommand _command;

        public RefreshCommand(ReparseCommand command) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _command = command;
        }        

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _command?.CanExecute(parameter) ?? false;
        }

        protected override void OnExecute(object parameter)
        {
            _command?.Execute(parameter);
        }
    }
}
