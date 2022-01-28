using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveCloserToUsage;
using Rubberduck.UI.Command;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.MoveCloserToUsage
{
    class MoveCloserToUsageViewModel : RefactoringViewModelBase<MoveCloserToUsageModel>
    {
        public MoveCloserToUsageViewModel(MoveCloserToUsageModel model) : base(model)
        {
            SetNewDeclarationStatementCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), (o) => SetNewDeclarationStatementExecute(o));
        }

        public Declaration Target => Model.Target;

        public string Instructions
        {
            get
            {
                if (Target == null)
                {
                    return RefactoringsUI.MoveCloserToUsageDialog_InstructionsLabelText;
                }

                return string.Format(RefactoringsUI.MoveCloserToUsageDialog_InstructionsLabelText, Target.IdentifierName);
            }
        }

        public DelegateCommand SetNewDeclarationStatementCommand { get; }

        void SetNewDeclarationStatementExecute(object param)
        {
            if (param is string newDeclarationStatement)
            {
                Model.DeclarationStatement = newDeclarationStatement;
            }

        }

    }
}
