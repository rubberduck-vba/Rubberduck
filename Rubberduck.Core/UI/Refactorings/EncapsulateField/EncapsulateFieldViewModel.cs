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
            PropertyName = model.TargetDeclaration.IdentifierName;
        }

        public Declaration TargetDeclaration
        {
            get => Model.TargetDeclaration;
            set
            {
                Model.TargetDeclaration = value;
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

        public bool CanHaveLet => Model.CanImplementLet;
        public bool CanHaveSet => Model.CanImplementSet;

        public bool IsLetSelected
        {
            get => Model.ImplementLetSetterType;
            set
            {
                Model.ImplementLetSetterType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public bool IsSetSelected
        {
            get => Model.ImplementSetSetterType;
            set
            {
                Model.ImplementSetSetterType = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public string PropertyName
        {
            get => Model.PropertyName;
            set
            {
                Model.PropertyName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidPropertyName));
                OnPropertyChanged(nameof(HasValidNames));
                OnPropertyChanged(nameof(PropertyPreview));
            }
        }

        public string ParameterName
        {
            get => Model.ParameterName;
            set
            {
                Model.ParameterName = value;
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