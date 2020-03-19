using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class MoveToFolderRefactoringFailedNotifier : RefactoringFailureNotifierBase
    {
        public MoveToFolderRefactoringFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => Resources.RubberduckUI.MoveToFolderDialog_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Module);
                default:
                    return base.Message(exception);
            }
        }
    }
}