using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public sealed class ReorderParametersDialog : RefactoringDialogBase<ReorderParametersModel, ReorderParametersView, ReorderParametersViewModel>
    {
        public ReorderParametersDialog(ReorderParametersViewModel vm) : base(vm)
        {
            Text = RubberduckUI.ReorderParamsDialog_Caption;
        }
    }
}
