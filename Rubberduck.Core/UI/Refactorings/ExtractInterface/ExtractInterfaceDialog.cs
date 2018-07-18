using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings
{
    internal sealed class ExtractInterfaceDialog : RefactoringDialogBase<ExtractInterfaceModel, ExtractInterfaceView, ExtractInterfaceViewModel>
    {
        protected override int MinHeight => 339;
        protected override int MinWidth => 459;

        private ExtractInterfaceDialog(ExtractInterfaceModel model, ExtractInterfaceView view, ExtractInterfaceViewModel viewModel) : base(model, view, viewModel)
        {
            Text = RubberduckUI.ExtractInterface_Caption;
        }
    }
}
