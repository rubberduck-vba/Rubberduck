using System;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Rename
{
    public class RenamePresenter
    {
        private readonly IRenameView _view;
        private readonly RenameModel _model;

        public RenamePresenter(IRenameView view, RenameModel model)
        {
            _view = view;
            _view.OkButtonClicked += OnViewOkButtonClicked;

            _model = model;
        }

        public RenameModel Show()
        {
            if (_model.Target != null)
            {
                _view.Target = _model.Target;
                _view.ShowDialog();
            }

            return _model;
        }

        public RenameModel Show(Declaration target)
        {
            _model.PromptIfTargetImplementsInterface(ref target);
            _model.Target = target;
            _view.Target = target;
            _view.ShowDialog();
            return _model;
        }

        private void OnViewOkButtonClicked(object sender, EventArgs e)
        {
            _model.NewName = _view.NewName;
        }
    }
}

