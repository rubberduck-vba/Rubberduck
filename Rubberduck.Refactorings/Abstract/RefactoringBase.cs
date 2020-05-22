using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings
{
    public abstract class RefactoringBase : IRefactoring
    {
        protected readonly ISelectionProvider SelectionProvider;

        protected RefactoringBase(ISelectionProvider selectionProvider)
        {
            SelectionProvider = selectionProvider;
        }

        public virtual void Refactor()
        {
            var activeSelection = SelectionProvider.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                throw new NoActiveSelectionException();
            }

            Refactor(activeSelection.Value);
        }

        public virtual void Refactor(QualifiedSelection targetSelection)
        {
            var target = FindTargetDeclaration(targetSelection);

            if (target == null)
            {
                throw new NoDeclarationForSelectionException(targetSelection);
            }

            Refactor(target);
        }

        protected abstract Declaration FindTargetDeclaration(QualifiedSelection targetSelection);
        public abstract void Refactor(Declaration target);
    }
}