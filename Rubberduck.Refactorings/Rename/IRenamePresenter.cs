using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenamePresenter : IRefactoringPresenter<RenameModel>
    {
        RenameModel Show(Declaration target);
        RenameModel Model { get; }
    }
}