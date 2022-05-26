using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.PromoteToParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.PromoteToParameter
{
    public class PromoteToParameterRefactoring : RefactoringBase
    {
        private readonly IRefactoringAction<PromoteToParameterModel> _refactoringAction;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IMessageBox _messageBox;

        public PromoteToParameterRefactoring(
            PromoteToParameterRefactoringAction refactoringAction, 
            IMessageBox messageBox, 
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
        :base(selectionProvider)
        {
            _refactoringAction = refactoringAction;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _messageBox = messageBox;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable)
            {
                return null;
            }

            return selectedDeclaration;
        }

        public override void Refactor(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new InvalidDeclarationTypeException(target);
            }

            if (!target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                throw new TargetDeclarationIsNotContainedInAMethodException(target);
            }

            PromoteVariable(target);
        }

        private void PromoteVariable(Declaration target)
        {
            if (!PromptIfMethodImplementsInterface(target))
            {
                return;
            }

            var model = Model(target);
            _refactoringAction.Refactor(model);
        }

        private PromoteToParameterModel Model(Declaration target)
        {
            var enclosingMember = _selectedDeclarationProvider.SelectedMember(target.QualifiedSelection);
            return new PromoteToParameterModel(target, enclosingMember);
        }

        private bool PromptIfMethodImplementsInterface(Declaration targetVariable)
        {
            var functionDeclaration = _selectedDeclarationProvider.SelectedMember(targetVariable.QualifiedSelection);

            if (functionDeclaration == null || !functionDeclaration.IsInterfaceImplementation)
            {
                return true;
            }

            var interfaceImplementation = functionDeclaration.InterfaceMemberImplemented;

            if (interfaceImplementation == null)
            {
                return true;
            }

            var message = string.Format(RefactoringsUI.PromoteToParameter_PromptIfTargetIsInterface,
                functionDeclaration.IdentifierName, interfaceImplementation.ComponentName,
                interfaceImplementation.IdentifierName);

            return _messageBox.Question(message, RefactoringsUI.PromoteToParameter_Caption);
        }
    }
}
