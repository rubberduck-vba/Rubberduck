namespace Rubberduck.Refactoring
{
    public interface IRefactoringPresenterFactory<out TPresenter>
    {
        TPresenter Create();
    }
}