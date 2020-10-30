using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveToFolder;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class MoveContainingFolderRefactoringFailedNotifier : RefactoringFailureNotifierBase
    {
        public MoveContainingFolderRefactoringFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => RefactoringsUI.MoveFoldersDialog_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(
                        RefactoringsUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Module);
                case NoTargetFolderException noTargetFolder:
                    return RefactoringsUI.RefactoringFailure_NoTargetFolder;
                default:
                    return base.Message(exception);
            }
        }
    }
}