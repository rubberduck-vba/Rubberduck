using System;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings
{
    public abstract class RefactoringPresenterBase<TModel, TDialog, TView, TViewModel> : IDisposable, IRefactoringPresenter<TModel, TDialog, TView, TViewModel> 
        where TModel : class
        where TView : class, IRefactoringView<TModel>
        where TViewModel : RefactoringViewModelBase<TModel>
        where TDialog : class, IRefactoringDialog<TModel, TView, TViewModel>
    {
        private readonly TDialog _dialog;
        private readonly IRefactoringDialogFactory _factory;

        protected RefactoringPresenterBase(TModel model, IRefactoringDialogFactory factory, TView view)
        {
            _factory = factory;
            var viewModel = _factory.CreateViewModel<TModel, TViewModel>(model);
            _dialog = _factory.CreateDialog<TModel, TView, TViewModel, TDialog>(model, view, viewModel);
        }

        public TModel Model => _dialog.ViewModel.Model;
        public TDialog Dialog => _dialog;
        public TViewModel ViewModel => _dialog.ViewModel;
        public virtual RefactoringDialogResult DialogResult { get; protected set; }

        public virtual TModel Show()
        {
            DialogResult = _dialog.ShowDialog();
            return DialogResult == RefactoringDialogResult.Execute ? _dialog.ViewModel.Model : null;
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
                _factory.ReleaseViewModel(_dialog.ViewModel);
                _factory.ReleaseDialog(_dialog);
                _dialog.Dispose();
            }
        }
    }
}
