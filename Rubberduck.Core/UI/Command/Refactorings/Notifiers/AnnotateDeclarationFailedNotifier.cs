using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Interaction;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Resources;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class AnnotateDeclarationFailedNotifier : RefactoringFailureNotifierBase
    {
        public AnnotateDeclarationFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        { }

        protected override string Caption => RefactoringsUI.AnnotateDeclarationDialog_Caption;

        protected override string Message(RefactoringException exception)
        {
            if (exception is InvalidDeclarationTypeException invalidTypeException)
            {
                Logger.Warn(invalidTypeException);
                return string.Format(
                    RefactoringsUI.RefactoringFailure_AnnotateDeclaration_InvalidType,
                    invalidTypeException.TargetDeclaration.DeclarationType.ToLocalizedString());
            }

            return base.Message(exception);
        }
    }
}