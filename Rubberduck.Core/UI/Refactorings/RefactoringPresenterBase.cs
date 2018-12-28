using System;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public abstract class RefactoringPresenterBase<TModel, TDialog, TView, TViewModel> : IDisposable, IRefactoringPresenter<TModel, TDialog, TView, TViewModel> 
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TViewModel : class, IRefactoringViewModel<TModel>
        where TDialog : class, IRefactoringDialog<TModel, TView, TViewModel>
    {
        private readonly IRefactoringDialogFactory _factory;

        protected RefactoringPresenterBase(TModel model, IRefactoringDialogFactory factory)
        {
            _factory = factory;
            var view = _factory.CreateView<TModel, TView>(model);
            var viewModel = _factory.CreateViewModel<TModel, TViewModel>(model);
            Dialog = _factory.CreateDialog<TModel, TView, TViewModel, TDialog>(model, view, viewModel);
        }

        public TDialog Dialog { get; }
        public TModel Model => Dialog.Model;
        public TViewModel ViewModel => Dialog.ViewModel;
        public virtual RefactoringDialogResult DialogResult { get; protected set; }

        public virtual TModel Show()
        {
            DialogResult = Dialog.ShowDialog();
            return DialogResult == RefactoringDialogResult.Execute ? Dialog.ViewModel.Model : null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _factory.ReleaseViewModel(Dialog.ViewModel);
                _factory.ReleaseView(Dialog.View);
                _factory.ReleaseDialog(Dialog);

                Dialog.Dispose();
            }
        }
    }
}
