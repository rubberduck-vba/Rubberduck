using NLog;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using System.Linq;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersViewModel : ViewModelBase
    {
        public RubberduckParserState State { get; }

        private List<Parameter> _parameters;
        public List<Parameter> Parameters
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

        public RemoveParametersViewModel(RubberduckParserState state)
        {
            State = state;
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
            RemoveParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RemoveParameter((Parameter)param));
            RestoreParameterCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => RestoreParameter((Parameter)param));
        }

        private void RemoveParameter(Parameter parameter)
        {
            if (parameter != null)
            {
                parameter.IsRemoved = true;
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        private void RestoreParameter(Parameter parameter)
        {
            if (parameter != null)
            {
                parameter.IsRemoved = false;
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);
        
        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public CommandBase RemoveParameterCommand { get; }
        public CommandBase RestoreParameterCommand { get; }
    }
}
