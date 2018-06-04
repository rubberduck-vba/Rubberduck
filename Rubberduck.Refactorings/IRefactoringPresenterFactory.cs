namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenterFactory
    {
        TPresenter Create<TPresenter>();
        void Release<TPresenter>(TPresenter presenter);
    }
}