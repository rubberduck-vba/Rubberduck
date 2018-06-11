using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public sealed class ReorderParametersDialog : RefactoringDialogBase<ReorderParametersModel, ReorderParametersView, ReorderParametersViewModel>
    {
        protected override int MinHeight => 395;
        protected override int MinWidth => 494;

        public ReorderParametersDialog(ReorderParametersModel model, ReorderParametersViewModel viewModel) : base(model, viewModel)
        {
            Text = RubberduckUI.ReorderParamsDialog_Caption;
        }
    }
}
