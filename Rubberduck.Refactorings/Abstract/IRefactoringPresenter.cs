namespace Rubberduck.Refactorings
{
    public interface IRefactoringPresenter<out TModel>
    where TModel : class, IRefactoringModel
    {
        TModel Show();
    }
}