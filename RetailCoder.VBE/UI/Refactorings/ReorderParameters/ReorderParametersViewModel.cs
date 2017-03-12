using System;
using System.Collections.ObjectModel;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class ReorderParametersViewModel : ViewModelBase
    {
        public RubberduckParserState State { get; }

        private ObservableCollection<Parameter> _parameters;
        public ObservableCollection<Parameter> Parameters
        {
            get { return _parameters; }
            set
            {
                _parameters = value;
                OnPropertyChanged();
            }
        }

        public string SignaturePreview
        {
            get
            {
                // if there is only one parameter, we remove it without displaying the UI; this gets called anyway as part of the initialization process
                if (Parameters == null)
                {
                    return string.Empty;
                }

                return string.Empty;
            }
        }

        public ReorderParametersViewModel(RubberduckParserState state)
        {
            State = state;
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
            MoveParameterUpCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => MoveParameterUp((Parameter)param));
            MoveParameterDownCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => MoveParameterDown((Parameter)param));
        }

        private void MoveParameterUp(Parameter parameter)
        {
            if (parameter != null)
            {
                var currentIndex = Parameters.IndexOf(parameter);
                Parameters.Move(currentIndex, currentIndex - 1);
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        private void MoveParameterDown(Parameter parameter)
        {
            if (parameter != null)
            {
                var currentIndex = Parameters.IndexOf(parameter);
                Parameters.Move(currentIndex, currentIndex + 1);
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);

        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public CommandBase MoveParameterUpCommand { get; }
        public CommandBase MoveParameterDownCommand { get; }
    }
}
