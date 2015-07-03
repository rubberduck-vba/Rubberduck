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
        private readonly IRemoveParametersView _view;
        private readonly RemoveParametersModel _model;

        public RemoveParametersPresenter(IRemoveParametersView view, RemoveParametersModel model)
        {
            _view = view;
            _model = model;
        }

        public RemoveParametersModel Show()
        {
            if (_model.Parameters.Count == 0)
            {
                var message = string.Format(RubberduckUI.RemovePresenter_NoParametersError, _model.TargetDeclaration.IdentifierName);
                MessageBox.Show(message, RubberduckUI.RemoveParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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
