using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenameView : IDialogView
    {
        Declaration Target { get; set; }
        string NewName { get; set; }

        void Hide();
    }
}