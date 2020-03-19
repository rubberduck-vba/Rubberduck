using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UI.Refactorings.MoveMember
{
    public class MoveMemberPresenter : RefactoringPresenterBase<MoveMemberModel>, IMoveMemberPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(Resources.RubberduckUI.MoveMember_Caption, minimumHeight: 450, minimumWidth: 800);

        public MoveMemberPresenter(MoveMemberModel model, IRefactoringDialogFactory dialogFactory)
            : base(DialogData, model, dialogFactory) { }

        public override MoveMemberModel Show()
        {
            return  Model.SelectedDeclarations.FirstOrDefault() == null ? null : base.Show();
        }
    }
}
