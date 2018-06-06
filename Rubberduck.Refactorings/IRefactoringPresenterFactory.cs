namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenterFactory
    {
        TPresenter Create<TPresenter, TModel>(TModel model)
            where TPresenter : class
            where TModel : class;
        void Release<TPresenter>(TPresenter presenter);
    }
}