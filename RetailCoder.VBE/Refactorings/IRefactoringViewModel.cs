namespace Rubberduck.Refactorings
{
    public interface IRefactoringViewModel<TModel>
    {
        TModel Model { get; set; }
    }
}
