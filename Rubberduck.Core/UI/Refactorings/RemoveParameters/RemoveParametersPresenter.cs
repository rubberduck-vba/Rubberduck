using Rubberduck.Interaction;
using Rubberduck.Resources;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.Refactorings;

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
                    base.Show();
                    return DialogResult != RefactoringDialogResult.Execute ? null : Model;
            }
        }
    }
}
