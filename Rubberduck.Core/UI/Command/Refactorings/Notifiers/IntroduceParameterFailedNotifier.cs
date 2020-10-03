using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.IntroduceParameter;
using Rubberduck.CodeAnalysis.Inspections.Extensions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class IntroduceParameterFailedNotifier : RefactoringFailureNotifierBase
    {
        public IntroduceParameterFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => Resources.RubberduckUI.IntroduceParameter_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case TargetDeclarationIsNotContainedInAMethodException targetNotInMethod:
                    Logger.Warn(targetNotInMethod);
                    return string.Format(Resources.RubberduckUI.IntroduceParameterFailed_TargetNotContainedInMethod,
                        targetNotInMethod.TargetDeclaration.QualifiedName);
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType.ToLocalizedString(),
                        DeclarationType.Variable.ToLocalizedString());
                default:
                    return base.Message(exception);
            }
        }
    }
}