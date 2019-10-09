using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class ExtractInterfaceFailedNotifier : RefactoringFailureNotifierBase
    {
        public ExtractInterfaceFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => Resources.RubberduckUI.ExtractInterface_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType, 
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.ClassModule);
                case SuspendParserFailureException suspendParserFailure:
                    Logger.Warn(suspendParserFailure);
                    return Resources.RubberduckUI.RefactoringFailure_SuspendParserFailure;
                default:
                    return base.Message(exception);
            }
        }
    }
}