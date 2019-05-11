using Rubberduck.Resources;
using Rubberduck.Interaction;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ReorderParameters;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenter : RefactoringPresenterBase<ReorderParametersModel>, IReorderParametersPresenter
    {
        private static readonly DialogData DialogData =
            DialogData.Create(RubberduckUI.ReorderParamsDialog_Caption, 395, 494);
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenter(ReorderParametersModel model,
            IRefactoringDialogFactory dialogFactory, IMessageBox messageBox) : 
            base(DialogData,  model, dialogFactory)
        {
            _messageBox = messageBox;
        }

        public override ReorderParametersModel Show()
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

            if (Model.Parameters.Count < 2)
            {
                var message = string.Format(RubberduckUI.ReorderPresenter_LessThanTwoParametersError, Model.TargetDeclaration.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.ReorderParamsDialog_TitleText);
                return null;
            }

            return base.Show();
        }

        private bool UserConfirmsInterfaceTarget(ReorderParametersModel model)
        {
            var message = string.Format(RubberduckUI.Refactoring_TargetIsInterfaceMemberImplementation,
                model.OriginalTarget.IdentifierName, Model.TargetDeclaration.ComponentName, model.TargetDeclaration.IdentifierName);
            return UserConfirmsNewTarget(message);
        }

        private bool UserConfirmsNewTarget(string message)
        {
            return _messageBox.ConfirmYesNo(message, RubberduckUI.ReorderParamsDialog_TitleText);
        }
    }
}
