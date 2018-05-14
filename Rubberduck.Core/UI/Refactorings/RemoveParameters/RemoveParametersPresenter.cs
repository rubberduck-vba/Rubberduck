using Rubberduck.Interaction;
using Rubberduck.Resources;
using Rubberduck.Refactorings.RemoveParameters;
using System.Windows.Forms;
using System.Linq;

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
                return _model;
            }

            _view.ViewModel.Parameters = _model.Parameters.Select(p => p.ToViewModel()).ToList();
            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }
            _model.Parameters = _view.ViewModel.Parameters.Where(m => m.IsRemoved).Select(vm => vm.ToModel()).ToList();
            return _model;
        }
    }
}
