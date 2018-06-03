using System.Windows.Forms;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public sealed class RemoveParametersDialog : RefactoringDialogBase<RemoveParametersModel, RemoveParametersView, RemoveParametersViewModel>
    {
        public RemoveParametersDialog(RemoveParametersViewModel viewModel) : base(viewModel)
        {
            Text = RubberduckUI.RemoveParamsDialog_Caption;
        }
    }
}
