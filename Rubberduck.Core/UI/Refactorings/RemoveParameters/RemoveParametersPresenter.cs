using Rubberduck.Interaction;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Resources;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenter : RefactoringPresenterBase<RemoveParametersModel>, IRemoveParametersPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.RemoveParamsDialog_Caption, 395, 494);
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenter(RemoveParametersModel model,
            IRefactoringDialogFactory dialogFactory, IMessageBox messageBox) : 
            base(DialogData,  model, dialogFactory)
        {
            _messageBox = messageBox;
        }

        public override RemoveParametersModel Show()
        {
            if (Model.TargetDeclaration == null)
            {
                return null;
            }

            if (Model.IsInterfaceMemberRefactoring
                && !UserConfirmsInterfaceTarget(Model))
            {
                throw new RefactoringAbortedException();
            }

            switch (Model.Parameters.Count)
            {
                case 0:
                    var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, Model.TargetDeclaration.IdentifierName);
                    _messageBox.NotifyWarn(message, RubberduckUI.RemoveParamsDialog_TitleText);
                    return null;
                case 1:
                    Model.RemoveParameters = Model.Parameters;
                    return Model;
                default:
                    return base.Show();
            }
        }

        private bool UserConfirmsInterfaceTarget(RemoveParametersModel model)
        {
            var message = string.Format(RubberduckUI.Refactoring_TargetIsInterfaceMemberImplementation,
                model.OriginalTarget.IdentifierName, Model.TargetDeclaration.ComponentName, model.TargetDeclaration.IdentifierName);
            return UserConfirmsNewTarget(message);
        }

        private bool UserConfirmsNewTarget(string message)
        {
            return _messageBox.ConfirmYesNo(message, RubberduckUI.RemoveParamsDialog_TitleText);
        }
    }
}
