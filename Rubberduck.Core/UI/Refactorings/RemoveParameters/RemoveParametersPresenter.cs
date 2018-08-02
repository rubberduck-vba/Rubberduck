using Rubberduck.Interaction;
using Rubberduck.Resources;
using Rubberduck.Refactorings.RemoveParameters;
using System.Linq;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    internal class RemoveParametersPresenter : RefactoringPresenterBase<RemoveParametersModel, RemoveParametersDialog, RemoveParametersView, RemoveParametersViewModel>, IRemoveParametersPresenter
    {
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenter(RemoveParametersModel model,
            IRefactoringDialogFactory dialogFactory, RemoveParametersView view, IMessageBox messageBox) : base(model, dialogFactory, view)
        {
            _messageBox = messageBox;
        }

        public override RemoveParametersModel Show()
        {
            if (Model.TargetDeclaration == null)
            {
                return null;
            }

            if (Model.Parameters.Count == 0)
            {
                var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, Model.TargetDeclaration.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.RemoveParamsDialog_TitleText);
                return null;
            }

            if (Model.Parameters.Count == 1)
            {
                return Model;
            }

            ViewModel.Parameters = Model.Parameters.Select(p => p.ToViewModel()).ToList();
            Show();
            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }
            Model.RemoveParameters = ViewModel.Parameters.Where(m => m.IsRemoved).Select(vm => vm.ToModel()).ToList();
            return Model;
        }
    }
}
