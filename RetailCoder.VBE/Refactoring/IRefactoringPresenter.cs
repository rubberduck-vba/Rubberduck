using Rubberduck.Refactoring.ExtractMethod;

namespace Rubberduck.Refactoring
{
    public interface IRefactoringPresenter
    {
        ExtractMethodModel Show();
    }
}
