using System.Windows.Forms;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public class ReorderParametersPresenter
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
