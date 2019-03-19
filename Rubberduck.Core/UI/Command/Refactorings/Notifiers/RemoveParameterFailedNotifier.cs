using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class RemoveParameterFailedNotifier : RefactoringFailureNotifierBase
    {
        public RemoveParameterFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => Resources.RubberduckUI.RemoveParamsDialog_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Parameter);
                default:
                    return base.Message(exception);
            }
        }
    }
}