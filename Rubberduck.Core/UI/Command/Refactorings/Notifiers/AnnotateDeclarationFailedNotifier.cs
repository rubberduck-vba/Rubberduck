using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Interaction;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Resources;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class AnnotateDeclarationFailedNotifier : RefactoringFailureNotifierBase
    {
        public AnnotateDeclarationFailedNotifier(IMessageBox messageBox) 
            : base(messageBox)
        { }

        protected override string Caption => RubberduckUI.AnnotateDeclarationDialog_Caption;

        protected override string Message(RefactoringException exception)
        {
            if (exception is InvalidDeclarationTypeException invalidTypeException)
            {
                Logger.Warn(invalidTypeException);
                return string.Format(
                    RubberduckUI.RefactoringFailure_AnnotateDeclaration_InvalidType,
                    invalidTypeException.TargetDeclaration.DeclarationType.ToLocalizedString());
            }

            return base.Message(exception);
        }
    }
}