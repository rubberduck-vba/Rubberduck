using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings
{
    public abstract class RefactoringBase : IRefactoring
    {
        protected readonly IRewritingManager RewritingManager;
        protected readonly ISelectionService SelectionService;

        protected RefactoringBase(IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            RewritingManager = rewritingManager;
            SelectionService = selectionService;
        }

        public virtual void Refactor()
        {
            var activeSelection = SelectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return;
            }

            Refactor(activeSelection.Value);
        }

        public abstract void Refactor(QualifiedSelection target);
        public abstract void Refactor(Declaration target);
    }
}