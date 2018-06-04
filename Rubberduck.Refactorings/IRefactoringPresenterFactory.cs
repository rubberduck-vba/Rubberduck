namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenterFactory<TPresenter>
    {
        TPresenter Create();
        void Release(TPresenter presenter);
    }
}