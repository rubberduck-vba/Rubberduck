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
        private readonly Func<TModel, IDisposalActionContainer<TPresenter>> PresenterFactory;
        protected TModel Model;

        protected InteractiveRefactoringBase(IRewritingManager rewritingManager, ISelectionService selectionService, IRefactoringPresenterFactory factory) 
        :base(rewritingManager, selectionService)
        {
            PresenterFactory = ((model) => DisposalActionContainer.Create(factory.Create<TPresenter, TModel>(model), factory.Release));
        }

        protected void Refactor(TModel initialModel)
        {
            Model = initialModel;

            var model = initialModel;
            if (model == null)
            {
                throw new InvalidRefactoringModelException();
            }

            using (var presenterContainer = PresenterFactory(model))
            {
                var presenter = presenterContainer.Value;
                if (presenter == null)
                {
                    throw new InvalidRefactoringPresenterException();
                }

                model = presenter.Show();
                if (model == null)
                {
                    throw new InvalidRefactoringModelException();
                }

                Model = model;

                RefactorImpl(model);
            }
        }

        protected abstract void RefactorImpl(TModel model);

        public abstract override void Refactor(QualifiedSelection target);
        public abstract override void Refactor(Declaration target);
    }
}