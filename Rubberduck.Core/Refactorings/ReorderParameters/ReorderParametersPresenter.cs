using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.ReorderParameters;

namespace Rubberduck.Refactorings.ReorderParameters
{
    public interface IReorderParametersPresenter
    {
        ReorderParametersModel Show();
    }

    public class ReorderParametersPresenter : IReorderParametersPresenter
    {
        private readonly IRefactoringDialog<ReorderParametersViewModel> _view;
        private readonly ReorderParametersModel _model;
        private readonly IMessageBox _messageBox;

        public ReorderParametersPresenter(IRefactoringDialog<ReorderParametersViewModel> view, ReorderParametersModel model, IMessageBox messageBox)
        {
            _view = view;
            _model = model;
            _messageBox = messageBox;
        }

        public ReorderParametersModel Show()
        {
            if (_model.TargetDeclaration == null) { return null; }

            if (_model.Parameters.Count < 2)
            {
                var message = string.Format(RubberduckUI.ReorderPresenter_LessThanTwoParametersError, _model.TargetDeclaration.IdentifierName);
                _messageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }

            _view.ViewModel.Parameters = new ObservableCollection<Parameter>(_model.Parameters);

            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            _model.Parameters = _view.ViewModel.Parameters.ToList();
            return _model;
        }
    }
}
