using System;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public abstract class RefactoringFailureNotifierBase : IRefactoringFailureNotifier
    {
        private readonly IMessageBox _messageBox;

        protected readonly Logger Logger; 

        protected RefactoringFailureNotifierBase(IMessageBox messageBox)
        {
            _messageBox = messageBox;

            Logger = LogManager.GetLogger(GetType().FullName);
        }

        protected abstract string Caption { get; }

        public void Notify(RefactoringException exception)
        {
            var message = $"{RefactoringsUI.RefactoringFailure_BaseMessage}{Environment.NewLine}{Environment.NewLine}{Message(exception)}";
            _messageBox.NotifyWarn(message, Caption);
        }

        protected virtual string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case NoActiveSelectionException noActiveSelection:
                    Logger.Error(noActiveSelection);
                    return RefactoringsUI.RefactoringFailure_NoActiveSelection; 
                case NoDeclarationForSelectionException noDeclarationForSelection:
                    Logger.Warn(noDeclarationForSelection);
                    return RefactoringsUI.RefactoringFailure_NoTargetDeclarationForSelection;
                case TargetDeclarationIsNullException targetNull:
                    Logger.Error(targetNull);
                    return RefactoringsUI.RefactoringFailure_TargetNull;
                case TargetDeclarationNotUserDefinedException targetBuiltIn:
                    return string.Format(RefactoringsUI.RefactoringFailure_TargetNotUserDefined, targetBuiltIn.TargetDeclaration.QualifiedName);
                case SuspendParserFailureException suspendParserFailure:
                    Logger.Warn(suspendParserFailure);
                    return RefactoringsUI.RefactoringFailure_SuspendParserFailure;
                case AffectedModuleIsStaleException affectedModuleIsStale:
                    return string.Format(
                        RefactoringsUI.RefactoringFailure_AffectedModuleIsStale,
                        affectedModuleIsStale.StaleModule.ToString());
                default:
                    Logger.Error(exception);
                    return string.Empty;
            }
        }
    }
}