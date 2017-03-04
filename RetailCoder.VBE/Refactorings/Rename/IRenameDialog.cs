using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenameDialog : IDialogView
    {
        Declaration Target { get; set; }
        string NewName { get; set; }
    }
}
