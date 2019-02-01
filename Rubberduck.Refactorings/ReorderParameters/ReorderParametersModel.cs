using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersModel
    {
        public RubberduckParserState State { get; }
        public IEnumerable<Declaration> Declarations { get; }

        public Declaration TargetDeclaration { get; private set; }
        public List<Parameter> Parameters { get; set; }

        public ReorderParametersModel(RubberduckParserState state, QualifiedSelection selection)
        {
            State = state;
            Declarations = state.AllUserDeclarations;
            
            AcquireTarget(selection);

            Parameters = new List<Parameter>();
            LoadParameters();
        }

        private void AcquireTarget(QualifiedSelection selection)
        {
            TargetDeclaration = Declarations.FindTarget(selection, ValidDeclarationTypes);
            TargetDeclaration = PromptIfTargetImplementsInterface();
            TargetDeclaration = GetEvent();
            TargetDeclaration = GetGetter();
        }

        private void LoadParameters()
        {
            if (TargetDeclaration == null) { return; }

            Parameters = ((IParameterizedDeclaration) TargetDeclaration).Parameters.Select((param, idx) => new Parameter(param, idx)).ToList();

            if (TargetDeclaration.DeclarationType == DeclarationType.PropertyLet ||
                TargetDeclaration.DeclarationType == DeclarationType.PropertySet)
            {
                Parameters.Remove(Parameters.Last());
            }
        }

        public static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private Declaration PromptIfTargetImplementsInterface()
        {
            if (!(TargetDeclaration is ModuleBodyElementDeclaration member) || !member.IsInterfaceImplementation)
            {
                return TargetDeclaration;
            }

            var interfaceMember = member.InterfaceMemberImplemented;
            var message = 
                string.Format(RubberduckUI.Refactoring_TargetIsInterfaceMemberImplementation, TargetDeclaration.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);
            var args = new RefactoringConfirmEventArgs(message) {Confirm = true};
            ConfirmReorderParameter?.Invoke(this, args);
            return args.Confirm ? interfaceMember : null;
        }

        public event EventHandler<RefactoringConfirmEventArgs> ConfirmReorderParameter;

        private Declaration GetEvent()
        {
            foreach (var events in Declarations.Where(item => item.DeclarationType == DeclarationType.Event))
            {
                if (Declarations.FindHandlersForEvent(events).Any(reference => Equals(reference.Item2, TargetDeclaration)))
                {
                    return events;
                }
            }
            return TargetDeclaration;
        }

        private Declaration GetGetter()
        {
            if (TargetDeclaration == null)
            {
                return null;
            }

            if (TargetDeclaration.DeclarationType != DeclarationType.PropertyLet &&
                TargetDeclaration.DeclarationType != DeclarationType.PropertySet)
            {
                return TargetDeclaration;
            }

            var getter = Declarations.FirstOrDefault(item => item.Scope == TargetDeclaration.Scope &&
                                          item.IdentifierName == TargetDeclaration.IdentifierName &&
                                          item.DeclarationType == DeclarationType.PropertyGet);

            return getter ?? TargetDeclaration;
        }
    }
}
