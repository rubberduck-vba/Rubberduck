using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings
{
    internal sealed class ExtractInterfaceDialog : RefactoringDialogBase<ExtractInterfaceModel, ExtractInterfaceView, ExtractInterfaceViewModel>
    {
        private ExtractInterfaceDialog(ExtractInterfaceViewModel viewModel) : base(viewModel)
        {
            Text = RubberduckUI.ExtractInterface_Caption;
        }
    }
}
