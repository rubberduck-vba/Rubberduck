using NLog;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersViewModel : ViewModelBase
    {
        public List<Parameter> Parameters { get; }

        public RemoveParametersViewModel(List<Parameter> parameters)
        {
            Parameters = parameters;

            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => SaveAndCloseWindow());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => CloseWindow());
            RemoveParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RemoveParameter((Parameter)param));
            RestoreParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RestoreParameter((Parameter)param));
        }

        private void RemoveParameter(Parameter parameter)
        {
            if (parameter != null)
            {
                parameter.IsRemoved = true;
            }
        }

        private void RestoreParameter(Parameter parameter)
        {
            if (parameter != null)
            {
                parameter.IsRemoved = false;
            }
        }

        public event EventHandler OnWindowClosed;
        private void CloseWindow() => OnWindowClosed?.Invoke(this, EventArgs.Empty);

        private void SaveAndCloseWindow() => CloseWindow();
        
        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public CommandBase RemoveParameterCommand { get; }
        public CommandBase RestoreParameterCommand { get; }
    }
}
