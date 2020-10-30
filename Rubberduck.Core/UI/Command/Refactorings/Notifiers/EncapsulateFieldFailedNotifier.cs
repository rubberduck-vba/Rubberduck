using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class EncapsulateFieldFailedNotifier : RefactoringFailureNotifierBase
    {
        public EncapsulateFieldFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => RefactoringsUI.EncapsulateField_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(RefactoringsUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Variable);
                default:
                    return base.Message(exception);
            }
        }
    }
}