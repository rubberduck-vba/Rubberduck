using NLog;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersViewModel : ViewModelBase
    {
        private List<Parameter> _parameters;
        public List<Parameter> Parameters
        {
            get { return _parameters; }
            set
            {
                _parameters = value;
                OnPropertyChanged();
            }
        }

        public RemoveParametersViewModel()
        {
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
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

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);
        
        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public CommandBase RemoveParameterCommand { get; }
        public CommandBase RestoreParameterCommand { get; }
    }
}
