using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Resources;
using Rubberduck.Interaction;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    internal class ReorderParametersPresenter : RefactoringPresenterBase<ReorderParametersModel, ReorderParametersDialog, ReorderParametersView, ReorderParametersViewModel>, IReorderParametersPresenter
    {
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenter(ReorderParametersModel model,
            IRefactoringDialogFactory dialogFactory, ReorderParametersView view, IMessageBox messageBox) : base(model, dialogFactory, view)
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

            base.Show();
            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }

            Model.Parameters = ViewModel.Parameters.ToList();
            return Model;
        }
    }
}
