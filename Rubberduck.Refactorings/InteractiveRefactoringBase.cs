using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using System;

namespace Rubberduck.Refactorings
{
    public abstract class InteractiveRefactoringBase<TPresenter, TModel> : RefactoringBase where TPresenter : class where TModel : class
    {
        protected readonly Func<TModel, IDisposalActionContainer<TPresenter>> PresenterFactory;

        public InteractiveRefactoringBase(IRewritingManager rewritingManager, ISelectionService selectionService, IRefactoringPresenterFactory factory) 
        :base(rewritingManager, selectionService)
        {
            PresenterFactory = ((model) => DisposalActionContainer.Create(factory.Create<TPresenter, TModel>(model), factory.Release));
        }

        public abstract override void Refactor(QualifiedSelection target);
        public abstract override void Refactor(Declaration target);
    }
}