using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.UI.Refactorings.Rename;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenamePresenter
    {
        RenameModel Show();
        RenameModel Show(Declaration target);
        RenameModel Model { get; }
    }

    public class RenamePresenter : IRenamePresenter
    {
        private readonly IRefactoringDialog<RenameViewModel> _view;
        private readonly RenameModel _model;

        public RenamePresenter(IRefactoringDialog<RenameViewModel> view, RenameModel model)
        {
            _view = view;

            _model = model;
        }

        public RenameModel Model => _model;

        public RenameModel Show()
        {
            if (_model.Target == null) { return null; }

            return Show(_model.Target);
        }

        public RenameModel Show(Declaration target)
        {
            if(null == target) { return null; }

            _model.Target = target;
            _view.ViewModel.Target = target;

            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            _model.NewName = _view.ViewModel.NewName;
            return _model;
        }
    }
}

