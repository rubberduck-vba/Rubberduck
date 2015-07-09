using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenamePresenter
    {
        RenameModel Show();
        RenameModel Show(Declaration target);
    }

    public class RenamePresenter : IRenamePresenter
    {
        private readonly IRenameView _view;
        private readonly RenameModel _model;

        public RenamePresenter(IRenameView view, RenameModel model)
        {
            _view = view;

            _model = model;
        }

        public RenameModel Show()
        {
            if (_model.Target == null) { return null; }

            _view.Target = _model.Target;

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.NewName = _view.NewName;
            return _model;
        }

        public RenameModel Show(Declaration target)
        {
            _model.PromptIfTargetImplementsInterface(ref target);
            _model.Target = target;
            _view.Target = target;

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            _model.NewName = _view.NewName;
            return _model;
        }
    }
}

