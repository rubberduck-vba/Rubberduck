using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.Rename
{
    internal class RenamePresenter : RefactoringPresenterBase<RenameModel>, IRenamePresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RubberduckUI.RenameDialog_Caption, 164, 684);

        public RenamePresenter(RenameModel model, IRefactoringDialogFactory dialogFactory) : 
            base(DialogData,  model, dialogFactory) { }

        public override RenameModel Show()
        {
            return Model.Target == null ? null : base.Show();
        }

        public RenameModel Show(Declaration target)
        {
            if (null == target)
            {
                return null;
            }

            Model.Target = target;

            var model = Show();

            if (DialogResult != RefactoringDialogResult.Execute)
            {
                return null;
            }
            
            return model;
        }
    }
}

