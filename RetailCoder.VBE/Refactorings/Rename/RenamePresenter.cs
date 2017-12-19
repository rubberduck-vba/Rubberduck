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

        public RenamePresenter(IRefactoringDialog<RenameViewModel> view, RenameModel model)
        {
            _view = view;

            Model = model;
        }

        public RenameModel Model { get; }

        public RenameModel Show()
        {
            return Model.Target == null ? null : Show(Model.Target);
        }

        public RenameModel Show(Declaration target)
        {
            if (null == target)
            {
                return null;
            }

            Model.Target = target;
            _view.ViewModel.Target = target;

            _view.ShowDialog();

            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            Model.NewName = _view.ViewModel.NewName;
            return Model;
        }
    }
}

