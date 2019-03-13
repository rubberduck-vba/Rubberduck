using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class ReorderParametersViewModel : RefactoringViewModelBase<ReorderParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ReorderParametersViewModel(IDeclarationFinderProvider declarationFinderProvider, ReorderParametersModel model) : base(model)
        {
            _declarationFinderProvider = declarationFinderProvider;
            MoveParameterUpCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => MoveParameterUp((Parameter)param));
            MoveParameterDownCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => MoveParameterDown((Parameter)param));
            _parameters = new ObservableCollection<Parameter>(model.Parameters);
        }

        private void UpdateModelParameters()
        {
            Model.Parameters = _parameters.ToList();
        }

        private ObservableCollection<Parameter> _parameters;
        public ObservableCollection<Parameter> Parameters
        {
            get => _parameters;
            set
            {
                _parameters = value;
                UpdateModelParameters();
                OnPropertyChanged();
            }
        }

        public string SignaturePreview
        {
            get
            {
                // if there is only one parameter, we reorder it without displaying the UI; this gets called anyway as part of the initialization process
                if (Parameters == null)
                {
                    return string.Empty;
                }

                var member = Parameters[0].Declaration.ParentDeclaration;

                if (member.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    var getter = _declarationFinderProvider.DeclarationFinder
                        .UserDeclarations(DeclarationType.PropertyGet)
                        .FirstOrDefault(item => item.Scope == member.Scope
                                                && item.IdentifierName == member.IdentifierName);

                    var letter = _declarationFinderProvider.DeclarationFinder
                        .UserDeclarations(DeclarationType.PropertyLet)
                        .FirstOrDefault(item => item.Scope == member.Scope
                                                && item.IdentifierName == member.IdentifierName);

                    var setter = _declarationFinderProvider.DeclarationFinder
                        .UserDeclarations(DeclarationType.PropertySet)
                        .FirstOrDefault(item => item.Scope == member.Scope
                                                && item.IdentifierName == member.IdentifierName);

                    var signature = string.Empty;
                    if (getter != null)
                    {
                        signature += GetSignature((PropertyGetDeclaration)getter);
                    }
                    if (letter != null)
                    {
                        if (!string.IsNullOrEmpty(signature)) { signature += Environment.NewLine; }
                        signature += GetSignature((PropertyLetDeclaration)letter);
                    }
                    if (setter != null)
                    {
                        if (!string.IsNullOrEmpty(signature)) { signature += Environment.NewLine; }
                        signature += GetSignature((PropertySetDeclaration)setter);
                    }

                    return signature;
                }

                return GetSignature((dynamic)member);
            }
        }

        private string GetSignature(SubroutineDeclaration member)
        {
            var signature = new StringBuilder();
            signature.Append(member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString());
            signature.Append($" Sub {member.IdentifierName}(");

            var selectedParams = Parameters.Select(s => s.Name);
            signature.Append($", {selectedParams})");
            return signature.ToString();
        }

        private string GetSignature(FunctionDeclaration member)
        {
            var signature = new StringBuilder();
            signature.Append(member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString());
            signature.Append($" Function {member.IdentifierName}(");

            var selectedParams = Parameters.Select(s => s.Name);
            signature.Append($", {selectedParams}) As {member.AsTypeName}");
            return signature.ToString();
        }

        private string GetSignature(EventDeclaration member)
        {
            var signature = new StringBuilder();
            signature.Append(member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString());
            signature.Append($" Event {member.IdentifierName}(");

            var selectedParams = Parameters.Select(s => s.Name);
            signature.Append($", {selectedParams})");
            return signature.ToString();
        }

        private string GetSignature(PropertyGetDeclaration member)
        {
            var signature = new StringBuilder();
            signature.Append(member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString());
            signature.Append($" Property Get {member.IdentifierName}(");

            var selectedParams = Parameters.Select(s => s.Name);
            signature.Append($", {selectedParams}) As {member.AsTypeName}");
            return signature.ToString();
        }

        private string GetSignature(PropertyLetDeclaration member)
        {
            var signature = new StringBuilder();
            signature.Append(member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString());
            signature.Append($" Property Let {member.IdentifierName}(");

            var selectedParams = Parameters.Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter((ParameterDeclaration)member.Parameters.Last(), -1).Name);
            signature.Append($", {selectedParams})");
            return signature.ToString();
        }

        private string GetSignature(PropertySetDeclaration member)
        {
            var signature = new StringBuilder();
            signature.Append(member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString());
            signature.Append($" Property Set {member.IdentifierName}(");

            var selectedParams = Parameters.Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter((ParameterDeclaration)member.Parameters.Last(), -1).Name);
            signature.Append($", {selectedParams})");
            return signature.ToString();
        }

        public void UpdatePreview() => OnPropertyChanged(nameof(SignaturePreview));

        private void MoveParameterUp(Parameter parameter)
        {
            if (parameter != null)
            {
                var currentIndex = Parameters.IndexOf(parameter);
                Parameters.Move(currentIndex, currentIndex - 1);
                UpdateModelParameters();
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        private void MoveParameterDown(Parameter parameter)
        {
            if (parameter != null)
            {
                var currentIndex = Parameters.IndexOf(parameter);
                Parameters.Move(currentIndex, currentIndex + 1);
                UpdateModelParameters();
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        protected override bool DialogOkPossible()
        {
            return IsValidParamOrder();
        }

        private bool IsValidParamOrder()
        {
            var indexOfFirstOptionalParam = Model.Parameters.FindIndex(param => param.IsOptional);
            if (indexOfFirstOptionalParam >= 0)
            {
                for (var index = indexOfFirstOptionalParam + 1; index < Model.Parameters.Count; index++)
                {
                    if (!Model.Parameters.ElementAt(index).IsOptional)
                    {
                        return false;
                    }
                }
            }

            var indexOfParamArray = Model.Parameters.FindIndex(param => param.IsParamArray);
            if (indexOfParamArray >= 0 && indexOfParamArray != Model.Parameters.Count - 1)
            {
                return false;
            }

            return true;
        }

        public CommandBase MoveParameterUpCommand { get; }
        public CommandBase MoveParameterDownCommand { get; }
    }
}
