using System.Windows.Forms;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public interface IRemoveParametersPresenter
    {
        RemoveParametersModel Show();
    }

    public class RemoveParametersPresenter : IRemoveParametersPresenter
    {
        private readonly IRemoveParametersDialog _view;
        private readonly RemoveParametersModel _model;
        private readonly IMessageBox _messageBox;

        public RemoveParametersPresenter(IRemoveParametersDialog view, RemoveParametersModel model, IMessageBox messageBox)
        {
            _view = view;
            _model = model;
            _messageBox = messageBox;
        }

        public RemoveParametersModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            if (_model.Parameters.Count == 0)
            {
                var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, _model.TargetDeclaration.IdentifierName);
                _messageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }

            if (_model.Parameters.Count == 1)
            {
                _model.Parameters[0].IsRemoved = true;
                return _model;
            }

            _view.Parameters = _model.Parameters;
            _view.InitializeParameterGrid();

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.Parameters = _view.Parameters;
            return _model;
        }
    }
}
