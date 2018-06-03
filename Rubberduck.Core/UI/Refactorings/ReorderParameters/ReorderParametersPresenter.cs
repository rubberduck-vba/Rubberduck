using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Resources;
using Rubberduck.Interaction;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.ReorderParameters;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenter : RefactoringPresenterBase<ReorderParametersModel, ReorderParametersDialog, ReorderParametersView, ReorderParametersViewModel>
    {
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenter(ReorderParametersModel model,
            IRefactoringDialogFactory<ReorderParametersModel, ReorderParametersView, ReorderParametersViewModel,
                ReorderParametersDialog> dialogFactory, IMessageBox messageBox) : base(model, dialogFactory)
        {
            _messageBox = messageBox;
        }

        public override ReorderParametersModel Show()
        {
            if (Model.TargetDeclaration == null) { return null; }

            if (Model.Parameters.Count < 2)
            {
                var message = string.Format(RubberduckUI.ReorderPresenter_LessThanTwoParametersError, Model.TargetDeclaration.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.ReorderParamsDialog_TitleText);
                return null;
            }

            ViewModel.Parameters = new ObservableCollection<Parameter>(Model.Parameters);

            Show();
            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }

            Model.Parameters = ViewModel.Parameters.ToList();
            return Model;
        }
    }
}
