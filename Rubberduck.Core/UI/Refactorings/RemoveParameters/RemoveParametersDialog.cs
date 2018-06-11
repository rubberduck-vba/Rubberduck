using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public sealed class RemoveParametersDialog : RefactoringDialogBase<RemoveParametersModel, RemoveParametersView, RemoveParametersViewModel>
    {
        protected override int MinHeight => 395;
        protected override int MinWidth => 494;

        public RemoveParametersDialog(RemoveParametersModel model, RemoveParametersViewModel viewModel) : base(model, viewModel)
        {
            Text = RubberduckUI.RemoveParamsDialog_Caption;
        }
    }
}
