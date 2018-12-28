using NLog;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersViewModel : RefactoringViewModelBase<RemoveParametersModel>
    {
        private readonly IMessageBox _messageBox;

        public RemoveParametersViewModel(RubberduckParserState state, RemoveParametersModel model, IMessageBox messageBox) : base(model)
        {
            State = state;
            RemoveParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RemoveParameter((ParameterViewModel)param));
            RestoreParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RestoreParameter((ParameterViewModel)param));
            _messageBox = messageBox;
            model.ConfirmRemoveParameter += ConfirmRemoveParameterHandler;

            Parameters = model.Parameters.Select(p => p.ToViewModel()).ToList();
        }

        private void ConfirmRemoveParameterHandler(object sender, RefactoringConfirmEventArgs e)
        {
            e.Confirm = _messageBox.ConfirmYesNo(e.Message, RubberduckUI.ReorderParamsDialog_TitleText);
        }

        public RubberduckParserState State { get; }

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
                    var declarations = State.AllUserDeclarations;

                    var getter = declarations.FirstOrDefault(item => item.Scope == member.Scope &&
                                                                     item.IdentifierName == member.IdentifierName &&
                                                                     item.DeclarationType == DeclarationType.PropertyGet);

                    var letter = declarations.FirstOrDefault(item => item.Scope == member.Scope &&
                                                                     item.IdentifierName == member.IdentifierName &&
                                                                     item.DeclarationType == DeclarationType.PropertyLet);

                    var setter = declarations.FirstOrDefault(item => item.Scope == member.Scope &&
                                                                     item.IdentifierName == member.IdentifierName &&
                                                                     item.DeclarationType == DeclarationType.PropertySet);

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
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Sub " + member.IdentifierName + "(";

            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        private string GetSignature(FunctionDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Function " + member.IdentifierName + "(";

            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ") As " + member.AsTypeName;
        }

        private string GetSignature(EventDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Event " + member.IdentifierName + "(";

            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        private string GetSignature(PropertyGetDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Property Get " + member.IdentifierName + "(";

            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ") As " + member.AsTypeName;
        }

        private string GetSignature(PropertyLetDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Property Let " + member.IdentifierName + "(";

            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter(member.Parameters.Last()).Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        private string GetSignature(PropertySetDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Property Set " + member.IdentifierName + "(";

            var selectedParams = Parameters.Where(p => !p.IsRemoved).Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter(member.Parameters.Last()).Name);
            return signature + string.Join(", ", selectedParams) + ")";
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
