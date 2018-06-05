namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenterFactory
    {
        TPresenter Create<TPresenter, TModel>(TModel model);
        void Release<TPresenter>(TPresenter presenter);
    }
}