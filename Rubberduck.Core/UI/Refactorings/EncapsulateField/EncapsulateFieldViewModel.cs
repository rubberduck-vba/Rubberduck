using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Refactorings.EncapsulateField
{
    public class EncapsulateFieldViewModel : RefactoringViewModelBase<EncapsulateFieldModel>
    {
        public RubberduckParserState State { get; }
        public IIndenter Indenter { get; }

        public EncapsulateFieldViewModel(EncapsulateFieldModel model, RubberduckParserState state, IIndenter indenter) : base(model)
        {
            State = state;
            Indenter = indenter;

            IsLetSelected = true;
            CanHaveLet = true;
        }

        private Declaration _targetDeclaration;
        public Declaration TargetDeclaration
        {
            get => _targetDeclaration;
            set
            {
                _targetDeclaration = value;
                PropertyName = value.IdentifierName;
            }
        }

        private bool _expansionState = true;
        public bool ExpansionState
        {
            get => _expansionState;
            set
            {
                _expansionState = value;
                OnPropertyChanged();
                OnExpansionStateChanged(value);
            }
        }

        private bool _canHaveLet;
        public bool CanHaveLet
        {
            get => _canHaveLet;
            set
            {
                _canHaveLet = value;
                OnPropertyChanged();
            }
        }

        private bool _canHaveSet;
        public bool CanHaveSet
        {
            get => _canHaveSet;
            set
            {
                _canHaveSet = value;
                OnPropertyChanged();
            }
        }

        private bool _isLetSelected;
        public bool IsLetSelected
        {
            get => _isLetSelected;
            set
            {
                _isLetSelected = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        private bool _isSetSelected;
        public bool IsSetSelected
        {
            get => _isSetSelected;
            set
            {
                _isSetSelected = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        private string _propertyName;
        public string PropertyName
        {
            get => _propertyName;
            set
            {
                _propertyName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidPropertyName));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        private string _parameterName = "value";
        public string ParameterName
        {
            get => _parameterName;
            set
            {
                _parameterName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidParameterName));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public bool IsValidPropertyName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return TargetDeclaration != null
                       && !PropertyName.Equals(TargetDeclaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
                       && !PropertyName.Equals(ParameterName, StringComparison.InvariantCultureIgnoreCase)
                       && char.IsLetter(PropertyName.FirstOrDefault())
                       && !tokenValues.Contains(PropertyName, StringComparer.InvariantCultureIgnoreCase)
                       && PropertyName.All(c => char.IsLetterOrDigit(c) || c == '_');
            }
        }

        public bool IsValidParameterName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

                return TargetDeclaration != null
                       && !ParameterName.Equals(TargetDeclaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
                       && !ParameterName.Equals(PropertyName, StringComparison.InvariantCultureIgnoreCase)
                       && char.IsLetter(ParameterName.FirstOrDefault())
                       && !tokenValues.Contains(ParameterName, StringComparer.InvariantCultureIgnoreCase)
                       && ParameterName.All(c => char.IsLetterOrDigit(c) || c == '_');
            }
        }

        public bool HasValidNames => IsValidPropertyName && IsValidParameterName;

        public string PropertyPreview
        {
            get
            {
                if (TargetDeclaration == null)
                {
                    return string.Empty;
                }

                var previewGenerator = new PropertyGenerator
                {
                    PropertyName = PropertyName,
                    AsTypeName = TargetDeclaration.AsTypeName,
                    BackingField = TargetDeclaration.IdentifierName,
                    ParameterName = ParameterName,
                    GenerateSetter = IsSetSelected,
                    GenerateLetter = IsLetSelected
                };

                var field = $"{Tokens.Private} {TargetDeclaration.IdentifierName} {Tokens.As} {TargetDeclaration.AsTypeName}{Environment.NewLine}{Environment.NewLine}";

                var propertyText = previewGenerator.AllPropertyCode.Insert(0, field).Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                return string.Join(Environment.NewLine, Indenter.Indent(propertyText, true));
            }
        }

        public event EventHandler<bool> ExpansionStateChanged;
        private void OnExpansionStateChanged(bool value) => ExpansionStateChanged?.Invoke(this, value);
    }
}