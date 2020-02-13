using System.Linq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.IntroduceField;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldRefactoring : RefactoringBase
    {
        private readonly IBaseRefactoring<IntroduceFieldModel> _baseRefactoring;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public IntroduceFieldRefactoring(
            IntroduceFieldBaseRefactoring baseRefactoring, 
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
        :base(rewritingManager, selectionProvider)
        {
            _baseRefactoring = baseRefactoring;
            _selectedDeclarationProvider = selectedDeclarationProvider;
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

            if (new[] { DeclarationType.ClassModule, DeclarationType.ProceduralModule }.Contains(target.ParentDeclaration.DeclarationType))
            {
                throw new TargetIsAlreadyAFieldException(target);
            }

            PromoteVariable(target);
        }

        private void PromoteVariable(Declaration target)
        {
            var model = Model(target);
            _baseRefactoring.Refactor(model);
        }

        private static IntroduceFieldModel Model(Declaration target)
        {
            return new IntroduceFieldModel(target);
        }
    }
}
