using System;
using System.Runtime.InteropServices;
using NLog;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class DelegateCommand : CommandBase
    {
        private readonly Predicate<object> _canExecute;
        private readonly Action<object> _execute;

        public DelegateCommand(ILogger logger, Action<object> execute, Predicate<object> canExecute = null) : base(logger)
        {
            _canExecute = canExecute;
            _execute = execute;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _canExecute == null || _canExecute.Invoke(parameter);
        }

        protected override void ExecuteImpl(object parameter)
        {
            _execute.Invoke(parameter);
        }
    }
}
