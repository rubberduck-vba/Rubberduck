namespace Rubberduck.Refactorings
{
    public interface IRefactoringViewModel<TModel> : IRefactoringViewModel
    {
        TModel Model { get; }
    }

    public interface IRefactoringViewModel
    {
    }
}
