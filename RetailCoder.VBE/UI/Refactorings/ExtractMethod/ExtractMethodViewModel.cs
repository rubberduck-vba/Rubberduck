using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VB6.Interop.VBIDE;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ExtractMethodViewModel : ViewModelBase, IRefactoringViewModel<ExtractMethodModel>
    {
        public ExtractMethodViewModel()
        {
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
        }

        private ObservableCollection<ExtractedParameter> _parameters;
        public ObservableCollection<ExtractedParameter> Parameters
        {
            get => _parameters;
            set
            {
                _parameters = value;
                OnPropertyChanged();
            }
        }

        public IEnumerable<string> ComponentNames =>
            Model.State.AllUserDeclarations
                .Where(d => d.ComponentName == Model.CodeModule.Name && (d.DeclarationType & DeclarationType.Member) == DeclarationType.Member)
                .Select(d => d.IdentifierName);

        public string NewMethodName
        {
            get => Model.NewMethodName;
            set
            {
                Model.NewMethodName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidMethodName));
                OnPropertyChanged(nameof(PreviewCode));
            }
        }
        
        public string SourceMethodName => Model.SourceMethodName;
        public string PreviewCaption => $@"Code Preview extracted from {SourceMethodName}";
        public string PreviewCode => Model.PreviewCode;

        public IEnumerable<ExtractedParameter> Inputs;
        public IEnumerable<ExtractedParameter> Outputs;
        public IEnumerable<ExtractedParameter> Locals;
        public IEnumerable<ExtractedParameter> ReturnValues;
        public string Preview;
        public Accessibility Accessibility;

        public bool IsValidMethodName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
                return !string.IsNullOrWhiteSpace(NewMethodName)
                       && char.IsLetter(NewMethodName.FirstOrDefault())
                       && !NewMethodName.Any(c => !char.IsLetterOrDigit(c) && c != '_')
                       && !ComponentNames.Contains(NewMethodName, StringComparer.InvariantCultureIgnoreCase)
                       && !tokenValues.Contains(NewMethodName, StringComparer.InvariantCultureIgnoreCase);
            }
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);
        
        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public ExtractMethodModel Model { get; set; }
    }
}
