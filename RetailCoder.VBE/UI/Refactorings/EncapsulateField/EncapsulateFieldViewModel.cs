using System;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class EncapsulateFieldViewModel : ViewModelBase
    {
        public RubberduckParserState State { get; }
        public IIndenter Indenter { get; }

        public Declaration TargetDeclaration { get; set; }

        public EncapsulateFieldViewModel(RubberduckParserState state, IIndenter indenter)
        {
            State = state;
            Indenter = indenter;

            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());

            IsLetSelected = true;
            CanHaveLet = true;
        }

        private bool _canHaveLet;
        public bool CanHaveLet
        {
            get { return _canHaveLet; }
            set
            {
                _canHaveLet = value;
                OnPropertyChanged();
            }
        }

        private bool _canHaveSet;
        public bool CanHaveSet
        {
            get { return _canHaveSet; }
            set
            {
                _canHaveSet = value;
                OnPropertyChanged();
            }
        }

        private bool _isLetSelected;
        public bool IsLetSelected
        {
            get { return _isLetSelected; }
            set
            {
                _isLetSelected = value;
                OnPropertyChanged();
            }
        }

        private bool _isSetSelected;
        public bool IsSetSelected
        {
            get { return _isSetSelected; }
            set
            {
                _isSetSelected = value;
                OnPropertyChanged();
            }
        }

        private string _propertyName;
        public string PropertyName
        {
            get { return _propertyName; }
            set
            {
                _propertyName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidPropertyName));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        private string _parameterName;
        public string ParameterName
        {
            get { return _parameterName; }
            set
            {
                _parameterName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidParameterName));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public bool IsValidPropertyName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return TargetDeclaration != null
                       && ParameterName.Equals(TargetDeclaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
                       && ParameterName.Equals(PropertyName, StringComparison.InvariantCultureIgnoreCase)
                       && !char.IsLetter(ParameterName.FirstOrDefault())
                       && tokenValues.Contains(ParameterName, StringComparer.InvariantCultureIgnoreCase)
                       && ParameterName.Any(c => !char.IsLetterOrDigit(c) && c != '_');
            }
        }

        public bool IsValidParameterName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return TargetDeclaration != null
                       && PropertyName.Equals(TargetDeclaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
                       && PropertyName.Equals(ParameterName, StringComparison.InvariantCultureIgnoreCase)
                       && !char.IsLetter(PropertyName.FirstOrDefault())
                       && tokenValues.Contains(PropertyName, StringComparer.InvariantCultureIgnoreCase)
                       && PropertyName.Any(c => !char.IsLetterOrDigit(c) && c != '_');
            }
        }

        public bool HasValidNames => IsValidPropertyName && IsValidParameterName;

        public string PropertyPreview => string.Empty;

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);

        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
    }
}