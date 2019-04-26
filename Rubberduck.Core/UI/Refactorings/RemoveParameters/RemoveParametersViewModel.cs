using NLog;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersViewModel : RefactoringViewModelBase<RemoveParametersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public RemoveParametersViewModel(IDeclarationFinderProvider declarationFinderProvider, RemoveParametersModel model) : base(model)
        {
            _declarationFinderProvider = declarationFinderProvider;
            RemoveParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RemoveParameter((ParameterViewModel)param));
            RestoreParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RestoreParameter((ParameterViewModel)param));

            Parameters = model.Parameters.Select(p => p.ToViewModel()).ToList();
        }

        private void UpdateModelParameters()
        {
            Model.RemoveParameters = Parameters.Where(m => m.IsRemoved).Select(vm => vm.ToModel()).ToList();
        }

        private List<ParameterViewModel> _parameters;
        public List<ParameterViewModel> Parameters
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
                // if there is only one parameter, we remove it without displaying the UI; this gets called anyway as part of the initialization process
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
                    if (getter != null) { signature += GetSignature((PropertyGetDeclaration)getter); }
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

                return GetSignature((dynamic) member);
            }
        }

        private string GetSignature(SubroutineDeclaration member)
        {
            var access = member.Accessibility.TokenString();
            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return $"{access} Sub {member.IdentifierName}({string.Join(", ", selectedParams)})";
        }

        private string GetSignature(FunctionDeclaration member)
        {
            var access = member.Accessibility.TokenString();
            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return $"{access} Function {member.IdentifierName}({string.Join(", ", selectedParams)}) As {member.AsTypeName}";
        }

        private string GetSignature(EventDeclaration member)
        {
            var access = member.Accessibility.TokenString();
            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return $"{access} Event {member.IdentifierName}({string.Join(", ", selectedParams)})";
        }

        private string GetSignature(PropertyGetDeclaration member)
        {
            var access = member.Accessibility.TokenString();
            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return $"{access} Property Get {member.IdentifierName}({string.Join(", ", selectedParams)}) As {member.AsTypeName}";
        }

        private string GetSignature(PropertyLetDeclaration member)
        {
            var access = member.Accessibility.TokenString();
            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter(member.Parameters.Last()).Name);
            return $"{access} Property Let {member.IdentifierName}({string.Join(", ", selectedParams)})";
        }

        private string GetSignature(PropertySetDeclaration member)
        {
            var access = member.Accessibility.TokenString();
            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter(member.Parameters.Last()).Name);
            return $"{access} Property Set {member.IdentifierName}({string.Join(", ", selectedParams)})";
        }

        private void RemoveParameter(ParameterViewModel parameter)
        {
            if (parameter != null)
            {
                parameter.IsRemoved = true;
                UpdateModelParameters();
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        private void RestoreParameter(ParameterViewModel parameter)
        {
            if (parameter != null)
            {
                parameter.IsRemoved = false;
                UpdateModelParameters();
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        public CommandBase RemoveParameterCommand { get; }
        public CommandBase RestoreParameterCommand { get; }
    }
}
