using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenamePresenter
    {
        RenameModel Show();
        RenameModel Show(Declaration target);
        RenameModel Model { get; }
    }
}