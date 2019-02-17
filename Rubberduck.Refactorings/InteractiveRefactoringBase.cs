using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using System;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.Refactorings
{
    public abstract class InteractiveRefactoringBase<TPresenter, TModel> : RefactoringBase 
        where TPresenter : class, IRefactoringPresenter<TModel>
        where TModel : class, IRefactoringModel
    {
        protected readonly Func<TModel, IDisposalActionContainer<TPresenter>> PresenterFactory;
        protected TModel Model;

        protected InteractiveRefactoringBase(IRewritingManager rewritingManager, ISelectionService selectionService, IRefactoringPresenterFactory factory) 
        :base(rewritingManager, selectionService)
        {
            PresenterFactory = ((model) => DisposalActionContainer.Create(factory.Create<TPresenter, TModel>(model), factory.Release));
        }

        protected void Refactor(TModel initialModel)
        {
            Model = initialModel;
            if (Model == null)
            {
                throw new InvalidRefactoringModelException();
            }

            using (var presenterContainer = PresenterFactory(Model))
            {
                var presenter = presenterContainer.Value;
                if (presenter == null)
                {
                    throw new InvalidRefactoringPresenterException();
                }

                Model = presenter.Show();
                if (Model == null)
                {
                    throw new InvalidRefactoringModelException();
                }

                RefactorImpl(presenter);
            }
        }

        protected abstract void RefactorImpl(TPresenter presenter);

        public abstract override void Refactor(QualifiedSelection target);
        public abstract override void Refactor(Declaration target);
    }
}