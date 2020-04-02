namespace Rubberduck.Refactorings
{
    public interface IRefactoringUserInteraction<TModel>
        where TModel : class, IRefactoringModel
    {
        TModel UserModifiedModel(TModel model);
    }
}