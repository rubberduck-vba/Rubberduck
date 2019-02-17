using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using System;

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

        public override void Refactor(Declaration target)
        {
            Model = InitializeModel(target);
            if (Model == null)
            {
                return;
            }

            using (var presenterContainer = PresenterFactory(Model))
            {
                var presenter = presenterContainer.Value;
                if (presenter == null)
                {
                    return;
                }

                Model = presenter.Show();
                if (Model == null)
                {
                    return;
                }

                RefactorImpl(presenter);
            }
        }

        protected abstract TModel InitializeModel(Declaration target);
        protected abstract void RefactorImpl(TPresenter presenter);

        public abstract override void Refactor(QualifiedSelection target);


    }
}