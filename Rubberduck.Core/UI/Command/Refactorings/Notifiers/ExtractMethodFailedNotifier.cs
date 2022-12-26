using Rubberduck.Interaction;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.ExtractMethod;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class ExtractMethodFailedNotifier : RefactoringFailureNotifierBase
    {
        public ExtractMethodFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => RefactoringsUI.ExtractMethod_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case UnableToMoveVariableDeclarationException unableToMoveVariableDeclaration:
                    Logger.Warn(unableToMoveVariableDeclaration);
                    return RefactoringsUI.ExtractMethod_UnableToMoveVariableDeclarationMessage; //TODO - improve this message to show declaration
                case InvalidTargetSelectionException invalidTargetSelection:
                    return string.Format(RefactoringsUI.ExtractMethod_InvalidSelectionMessage, invalidTargetSelection.Message);
                //case 
                default:
                    return base.Message(exception);
            }
        }
    }
}