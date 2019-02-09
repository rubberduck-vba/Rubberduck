namespace Rubberduck.Refactorings
{
    public interface IRefactoringView
    {
        object DataContext { get; set; }
    }

    public interface IRefactoringView<TModel> : IRefactoringView
        where TModel : class
    { }
}
