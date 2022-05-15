using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveCloserToUsage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.MoveCloserToUsage
{
    class MoveCloserToUsagePresenter : RefactoringPresenterBase<MoveCloserToUsageModel>, IMoveCloserToUsagePresenter
    {

        private static readonly DialogData DialogData = DialogData.Create("RefactoringsUI.MoveCloserToUsageDialog_Caption", 164, 684);

        public MoveCloserToUsagePresenter(MoveCloserToUsageModel model, IRefactoringDialogFactory dialogFactory) : 
            base(DialogData, model, dialogFactory)
        {
        }

        public MoveCloserToUsageModel Show(VariableDeclaration target)
        {
            if (null == target)
            {
                return null;
            }

            Model.Target = target;

            return Show();
        }
    }
}
