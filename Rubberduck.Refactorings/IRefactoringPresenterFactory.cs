namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenterFactory
    {
        TPresenter Create<TPresenter, TModel>(TModel model)
            where TPresenter : class, IRefactoringPresenter<TModel>
            where TModel : class, IRefactoringModel;
        void Release<TPresenter>(TPresenter presenter);
    }
}