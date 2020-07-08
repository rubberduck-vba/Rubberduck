using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveCloserToUsage;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class MoveCloserToUsageFailedNotifier : RefactoringFailureNotifierBase
    {
        public MoveCloserToUsageFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        {}

        protected override string Caption => Resources.RubberduckUI.MoveCloserToUsage_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case TargetDeclarationConflictsWithPreexistingDeclaration conflictWithPreexistingDeclaration:
                    Logger.Warn(conflictWithPreexistingDeclaration);
                    return string.Format(Resources.RubberduckUI.MoveCloserToUsageFailure_ReferencingMethodHasSameNameDeclarationInScope,
                        conflictWithPreexistingDeclaration.TargetDeclaration.QualifiedName,
                        conflictWithPreexistingDeclaration.ConflictingDeclaration.QualifiedName);
                case TargetDeclarationNonPrivateInNonStandardModule nonPrivateInNonStandardModule:
                    Logger.Warn(nonPrivateInNonStandardModule);
                    return string.Format(Resources.RubberduckUI.MoveCloserToUsageFailure_TargetIsNonPrivateInNonStandardModule,
                        nonPrivateInNonStandardModule.TargetDeclaration.QualifiedName);
                case TargetDeclarationInDifferentNonStandardModuleException inDifferentNonStandardModule:
                    Logger.Warn(inDifferentNonStandardModule);
                    return string.Format(Resources.RubberduckUI.MoveCloserToUsageFailure_TargetIsInOtherNonStandardModule,
                        inDifferentNonStandardModule.TargetDeclaration.QualifiedName);
                case TargetDeclarationInDifferentProjectThanUses inDifferentProject:
                    Logger.Warn(inDifferentProject);
                    return string.Format(Resources.RubberduckUI.MoveCloserToUsageFailure_TargetIsInDifferentProject,
                        inDifferentProject.TargetDeclaration.QualifiedName);
                case TargetDeclarationUsedInMultipleMethodsException usedInMultiple:
                    Logger.Warn(usedInMultiple);
                    return string.Format(Resources.RubberduckUI.MoveCloserToUsageFailure_TargetIsUsedInMultipleMethods,
                        usedInMultiple.TargetDeclaration.QualifiedName);
                case TargetDeclarationNotUserDefinedException notUsed:
                    Logger.Warn(notUsed);
                    return string.Format(Resources.RubberduckUI.MoveCloserToUsageFailure_TargetHasNoReferences,
                        notUsed.TargetDeclaration.QualifiedName);
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Variable);
                default:
                    return base.Message(exception);
            }
        }
    }
}
