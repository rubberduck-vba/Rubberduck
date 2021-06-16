using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.ImplementInterface;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class ImplementInterfaceFailedNotifier : RefactoringFailureNotifierBase
    {
        public ImplementInterfaceFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => RefactoringsUI.ImplementInterface_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case NoImplementsStatementSelectedException noImplementsStatementSelected:
                    Logger.Warn(noImplementsStatementSelected);
                    return RefactoringsUI.ImplementInterfaceFailed_NoImplementsStatementSelected;
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(RefactoringsUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.ClassModule);
                default:
                    return base.Message(exception);
            }
        }
    }
}