namespace Rubberduck.Refactorings
{
    // TODO: This factory interface is not really needed and should be refactored out
    public interface IRefactoringPresenterFactory<out TPresenter>
    {
        TPresenter Create();
    }
}