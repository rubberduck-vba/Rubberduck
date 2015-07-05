using System.Windows.Forms;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public interface IReorderParametersPresenter
    {
        ReorderParametersModel Show();
    }

    public class ReorderParametersPresenter : IReorderParametersPresenter
    {
        private readonly IReorderParametersView _view;
        private readonly ReorderParametersModel _model;

        public ReorderParametersPresenter(IReorderParametersView view, ReorderParametersModel model)
        {
            _view = view;
            _model = model;
        }

        public ReorderParametersModel Show()
        {
            if (_model.Parameters.Count < 2)
            {
                var message = string.Format(RubberduckUI.ReorderPresenter_LessThanTwoParametersError, _model.TargetDeclaration.IdentifierName);
                MessageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
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
