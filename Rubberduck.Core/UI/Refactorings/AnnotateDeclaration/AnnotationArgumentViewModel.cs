using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    internal interface IAnnotationArgumentViewModel : INotifyPropertyChanged, INotifyDataErrorInfo
    {
        TypedAnnotationArgument Model { get; }
        IReadOnlyList<AnnotationArgumentType> ApplicableArgumentTypes { get; }
        IReadOnlyList<string> InspectionNames { get; }
        IReadOnlyList<string> BooleanValues { get; }
        bool CanEditArgumentType { get; }
        AnnotationArgumentType ArgumentType { get; set; }
        string ArgumentValue { get; set; }
    }

    internal class AnnotationArgumentViewModel : ViewModelBase, IAnnotationArgumentViewModel
    {
        private const int MaxAllowedCharacters = 511;

        private TypedAnnotationArgument _model;

        public AnnotationArgumentViewModel(TypedAnnotationArgument model, IReadOnlyList<string> inspectionNames)
        {
            _model = model;

            ApplicableArgumentTypes = ApplicableTypes(_model.ArgumentType);
            InspectionNames = inspectionNames;
            BooleanValues = new List<string> { "True", "False" };

            _model.ArgumentType = ApplicableArgumentTypes.FirstOrDefault();
            _model.Argument = string.IsNullOrEmpty(_model.Argument)
                ? DefaultValue(_model.ArgumentType)
                : _model.Argument;

            ValidateArgument();
        }

        private IReadOnlyList<AnnotationArgumentType> ApplicableTypes(AnnotationArgumentType argumentType)
        {
            return Enum.GetValues(typeof(AnnotationArgumentType))
                .OfType<AnnotationArgumentType>()
                .Where(t => argumentType.HasFlag(t))
                .ToList();
        }

        public TypedAnnotationArgument Model => _model;
        public IReadOnlyList<AnnotationArgumentType> ApplicableArgumentTypes { get; }

        public bool CanEditArgumentType => ApplicableArgumentTypes.Count > 1;
        public IReadOnlyList<string> InspectionNames { get; }
        public IReadOnlyList<string> BooleanValues { get; }

        public AnnotationArgumentType ArgumentType
        {
            get => _model.ArgumentType;
            set
            {
                if (_model.ArgumentType == value)
                {
                    return;
                }

                _model.ArgumentType = value;
                ArgumentValue = DefaultValue(value);
                ValidateArgument();
                OnPropertyChanged();
            }
        }

        private string DefaultValue(AnnotationArgumentType argumentType)
        {
            switch (argumentType)
            {
                case AnnotationArgumentType.Boolean:
                    return "True";
                case AnnotationArgumentType.Inspection:
                    return InspectionNames.FirstOrDefault() ?? string.Empty;
                default:
                    return string.Empty;
            }
        }

        public string ArgumentValue
        {
            get => _model.Argument;
            set
            {
                if (_model.Argument == value)
                {
                    return;
                }

                _model.Argument = value;
                ValidateArgument();
                OnPropertyChanged();
            }
        }

        private void ValidateArgument()
        {
            var errors = ArgumentValidationErrors();

            if (errors.Any())
            {
                SetErrors(nameof(ArgumentValue), errors);
            }
            else
            {
                ClearErrors();
            }
        }

        private List<string> ArgumentValidationErrors()
        {
            var errors = new List<string>();

            if (string.IsNullOrEmpty(ArgumentValue))
            {
                errors.Add(RubberduckUI.AnnotationArgument_ValidationError_EmptyArgument);
            }
            if (ArgumentValue.Length > MaxAllowedCharacters)
            {
                errors.Add(string.Format(RubberduckUI.AnnotationArgument_ValidationError_TooLong, MaxAllowedCharacters));
            }
            if (ContainsNewline(ArgumentValue))
            {
                errors.Add(RubberduckUI.AnnotationArgument_ValidationError_Newline);
            }
            else if (ContainsControlCharacter(ArgumentValue))
            {
                errors.Add(RubberduckUI.AnnotationArgument_ValidationError_SpecialCharacters);
            }

            switch (ArgumentType)
            {
                case AnnotationArgumentType.Attribute:
                    if (!ArgumentValue.StartsWith("VB_") || ContainsWhitespace(ArgumentValue))
                    {
                        errors.Add(RubberduckUI.AnnotationArgument_ValidationError_AttributeNameStart);
                    }
                    if (ContainsWhitespace(ArgumentValue))
                    {
                        errors.Add(RubberduckUI.AnnotationArgument_ValidationError_WhitespaceInAttribute);
                    }
                    break;
                case AnnotationArgumentType.Inspection:
                    if (!InspectionNames.Contains(ArgumentValue))
                    {
                        errors.Add(RubberduckUI.AnnotationArgument_ValidationError_InspectionName);
                    }
                    break;
                case AnnotationArgumentType.Boolean:
                    if (!bool.TryParse(ArgumentValue, out _))
                    {
                        errors.Add(RubberduckUI.AnnotationArgument_ValidationError_NotABoolean);
                    }
                    break;
                case AnnotationArgumentType.Number:
                    if (!decimal.TryParse(ArgumentValue, out _))
                    {
                        errors.Add(RubberduckUI.AnnotationArgument_ValidationError_NotANumber);
                    }
                    break;
            }

            return errors;
        }

        private bool ContainsNewline(string argumentText)
        {
            return argumentText.Contains('\n') || argumentText.Contains('\r');
        }

        private bool ContainsControlCharacter(string argumentText)
        {
            return argumentText.Any(char.IsControl);
        }

        private bool ContainsWhitespace(string argumentText)
        {
            return argumentText.Any(char.IsWhiteSpace);
        }
    }
}