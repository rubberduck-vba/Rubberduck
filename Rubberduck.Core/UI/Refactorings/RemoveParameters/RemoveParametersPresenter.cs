using Rubberduck.Interaction;
using Rubberduck.Resources;
using Rubberduck.UI.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class RemoveParametersPresenter : IRemoveParametersPresenter
    {
        private readonly IRefactoringDialog<RemoveParametersViewModel> _view;
        private readonly RemoveParametersModel _model;
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenter(IRefactoringDialog<RemoveParametersViewModel> view, RemoveParametersModel model, IMessageBox messageBox)
        {
            _view = view;
            _model = model;
            _messageBox = messageBox;
        }

        public RemoveParametersModel Show()
        {
            if (_model.TargetDeclaration == null)
            {
                return null;
            }

            if (_model.Parameters.Count == 0)
            {
                var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, _model.TargetDeclaration.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.RemoveParamsDialog_TitleText);
                return null;
            }

            if (_model.Parameters.Count == 1)
            {
                _model.Parameters[0].IsRemoved = true;
                return _model;
            }

            _view.ViewModel.Parameters = _model.Parameters;
            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            _model.Parameters = _view.ViewModel.Parameters;
            return _model;
        }
    }
}
