using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.Rename
{
    public sealed class RenameDialog : RefactoringDialogBase<RenameModel, RenameView, RenameViewModel>
    {
        protected override int MinHeight => 164;
        protected override int MinWidth => 684;

        public RenameDialog(RenameModel model, RenameViewModel viewModel) : base(model, viewModel)
        {
            Text = RubberduckUI.RenameDialog_Caption;
        }
    }
}
