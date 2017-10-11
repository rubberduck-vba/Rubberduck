using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings
{
    public class ExtractMethodViewModel : ViewModelBase
    {
        public ExtractMethodViewModel()
        {
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
        }

        private List<ExtractedParameter> _parameters;
        public List<ExtractedParameter> Parameters
        {
            get { return _parameters; }
            set
            {
                _parameters = value;
                OnPropertyChanged();
            }
        }

        public List<string> ComponentNames { get; set; }

        private string _methodName;
        public string MethodName
        {
            get { return _methodName; }
            set
            {
                _methodName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidMethodName));
            }
        }

        private string _oldMethodName;
        public string OldMethodName
        {
            get { return _oldMethodName; }
            set
            {
                if(string.IsNullOrWhiteSpace( _oldMethodName))
                {
                    _oldMethodName = value;
                }
            }
        }

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

                return !ComponentNames.Contains(MethodName)
                       && MethodName.Length > 1
                       && char.IsLetter(MethodName.FirstOrDefault())
                       && !tokenValues.Contains(MethodName, StringComparer.InvariantCultureIgnoreCase)
                       && !MethodName.Any(c => !char.IsLetterOrDigit(c) && c != '_');
            }
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);
        
        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
    }
}
