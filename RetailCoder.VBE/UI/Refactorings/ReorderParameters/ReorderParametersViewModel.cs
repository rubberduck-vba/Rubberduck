using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class ReorderParametersViewModel : ViewModelBase
    {
        public RubberduckParserState State { get; }

        private ObservableCollection<Parameter> _parameters;
        public ObservableCollection<Parameter> Parameters
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

                return GetSignature((dynamic)member);
            }
        }

        private string GetSignature(SubroutineDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Sub " + member.IdentifierName + "(";

            var selectedParams = Parameters.Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        private string GetSignature(FunctionDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Function " + member.IdentifierName + "(";

            var selectedParams = Parameters.Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ") As " + member.AsTypeName;
        }

        private string GetSignature(EventDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Event " + member.IdentifierName + "(";

            var selectedParams = Parameters.Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        private string GetSignature(PropertyGetDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Property Get " + member.IdentifierName + "(";

            var selectedParams = Parameters.Select(s => s.Name);
            return signature + string.Join(", ", selectedParams) + ") As " + member.AsTypeName;
        }

        private string GetSignature(PropertyLetDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Property Let " + member.IdentifierName + "(";

            var selectedParams = Parameters.Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter((ParameterDeclaration)member.Parameters.Last(), -1).Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        private string GetSignature(PropertySetDeclaration member)
        {
            var signature = member.Accessibility == Accessibility.Implicit ? string.Empty : member.Accessibility.ToString();
            signature += " Property Set " + member.IdentifierName + "(";

            var selectedParams = Parameters.Select(s => s.Name).ToList();
            selectedParams.Add(new Parameter((ParameterDeclaration)member.Parameters.Last(), -1).Name);
            return signature + string.Join(", ", selectedParams) + ")";
        }

        public ReorderParametersViewModel(RubberduckParserState state)
        {
            State = state;
            OkButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogOk());
            CancelButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => DialogCancel());
            MoveParameterUpCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => MoveParameterUp((Parameter)param));
            MoveParameterDownCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), param => MoveParameterDown((Parameter)param));
        }

        public void UpdatePreview() => OnPropertyChanged(nameof(SignaturePreview));

        private void MoveParameterUp(Parameter parameter)
        {
            if (parameter != null)
            {
                var currentIndex = Parameters.IndexOf(parameter);
                Parameters.Move(currentIndex, currentIndex - 1);
                
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        private void MoveParameterDown(Parameter parameter)
        {
            if (parameter != null)
            {
                var currentIndex = Parameters.IndexOf(parameter);
                Parameters.Move(currentIndex, currentIndex + 1);
                
                OnPropertyChanged(nameof(SignaturePreview));
            }
        }

        public event EventHandler<DialogResult> OnWindowClosed;
        private void DialogCancel() => OnWindowClosed?.Invoke(this, DialogResult.Cancel);
        private void DialogOk() => OnWindowClosed?.Invoke(this, DialogResult.OK);

        public CommandBase OkButtonCommand { get; }
        public CommandBase CancelButtonCommand { get; }
        public CommandBase MoveParameterUpCommand { get; }
        public CommandBase MoveParameterDownCommand { get; }
    }
}
