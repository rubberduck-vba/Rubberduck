using Rubberduck.Interaction;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.Rename;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class RenameFailedNotifier : RefactoringFailureNotifierBase
    {
        public RenameFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        {}

        protected override string Caption => Resources.RubberduckUI.RenameDialog_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case CodeModuleNotFoundException codeModuleNotFound:
                    Logger.Warn(codeModuleNotFound);
                    return string.Format(Resources.RubberduckUI.RenameFailure_TargetModuleWithoutCodeModule, codeModuleNotFound.TargetDeclaration.QualifiedModuleName);
                case TargetControlNotFoundException controlNotFound:
                    Logger.Warn(controlNotFound);
                    return string.Format(Resources.RubberduckUI.RenameFailure_TargetControlNotFound, controlNotFound.TargetDeclaration.QualifiedName);
                case TargetDeclarationIsStandardEventHandlerException standardHandler:
                    return string.Format(Resources.RubberduckUI.RenameFailure_StandardEventHandler, standardHandler.TargetDeclaration.QualifiedName);
                default:
                    return base.Message(exception);
            }
        }
    }
}