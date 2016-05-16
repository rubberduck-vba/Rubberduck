namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenterFactory<out TPresenter>
    {
        TPresenter Create();
    }
}
