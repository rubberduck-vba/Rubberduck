using System;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings
{
    public class RefactoringUserInteraction<TPresenter, TModel> : IRefactoringUserInteraction<TModel>
        where TPresenter : class, IRefactoringPresenter<TModel>
        where TModel : class, IRefactoringModel
    {
        private readonly Func<TModel, IDisposalActionContainer<TPresenter>> _presenterFactory;

        public RefactoringUserInteraction(IRefactoringPresenterFactory factory, IUiDispatcher uiDispatcher)
        {
            Action<TPresenter> presenterDisposalAction = (TPresenter presenter) => uiDispatcher.Invoke(() => factory.Release(presenter));
            _presenterFactory = ((model) => DisposalActionContainer.Create(factory.Create<TPresenter, TModel>(model), presenterDisposalAction));
        }

        public TModel UserModifiedModel(TModel model)
        {
            using (var presenterContainer = _presenterFactory(model))
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
            }

            return model;
        }
    }
}