using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class MoveMemberFailedNotifier : RefactoringFailureNotifierBase
    {
        public MoveMemberFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => Resources.RubberduckUI.MoveMember_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        $"{DeclarationType.Variable}, {DeclarationType.Constant}, {DeclarationType.Member}");
                case MoveMemberUnsupportedMoveException unsupportedMove:
                    Logger.Warn(unsupportedMove);
                    return string.Format( Resources.RubberduckUI.MoveMember_UnsupportedMoveExceptionFormat, unsupportedMove.TargetDeclaration.IdentifierName);
                default:
                    return base.Message(exception);
            }
        }
    }
}