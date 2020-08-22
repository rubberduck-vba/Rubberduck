using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveToFolder;

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
                    return string.Format(
                        Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Module);
                case NoTargetFolderException noTargetFolder:
                    return Resources.RubberduckUI.RefactoringFailure_NoTargetFolder;
                default:
                    return base.Message(exception);
            }
        }
    }
}