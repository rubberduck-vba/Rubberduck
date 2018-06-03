using System.Windows.Forms;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.Rename
{
    public sealed class RenameDialog : RefactoringDialogBase<RenameModel, RenameView, RenameViewModel>
    {
        public RenameDialog(RenameViewModel vm) : base(vm)
        {
            Text = RubberduckUI.RenameDialog_Caption;
        }
    }
}
