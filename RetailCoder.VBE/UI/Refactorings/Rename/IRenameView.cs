using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.Rename
{
    public interface IRenameView : IDialogView
    {
        Declaration Target { get; set; }
        string NewName { get; set; }
    }
}