using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveCloserToUsage;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class MoveCloserToUsageFailedNotifier : RefactoringFailureNotifierBase
    {
        public MoveCloserToUsageFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        {}

        protected override string Caption => RefactoringsUI.MoveCloserToUsage_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case TargetDeclarationConflictsWithPreexistingDeclaration conflictWithPreexistingDeclaration:
                    Logger.Warn(conflictWithPreexistingDeclaration);
                    return string.Format(RefactoringsUI.MoveCloserToUsageFailure_ReferencingMethodHasSameNameDeclarationInScope,
                        conflictWithPreexistingDeclaration.TargetDeclaration.QualifiedName,
                        conflictWithPreexistingDeclaration.ConflictingDeclaration.QualifiedName);
                case TargetDeclarationNonPrivateInNonStandardModule nonPrivateInNonStandardModule:
                    Logger.Warn(nonPrivateInNonStandardModule);
                    return string.Format(RefactoringsUI.MoveCloserToUsageFailure_TargetIsNonPrivateInNonStandardModule,
                        nonPrivateInNonStandardModule.TargetDeclaration.QualifiedName);
                case TargetDeclarationInDifferentNonStandardModuleException inDifferentNonStandardModule:
                    Logger.Warn(inDifferentNonStandardModule);
                    return string.Format(RefactoringsUI.MoveCloserToUsageFailure_TargetIsInOtherNonStandardModule,
                        inDifferentNonStandardModule.TargetDeclaration.QualifiedName);
                case TargetDeclarationInDifferentProjectThanUses inDifferentProject:
                    Logger.Warn(inDifferentProject);
                    return string.Format(RefactoringsUI.MoveCloserToUsageFailure_TargetIsInDifferentProject,
                        inDifferentProject.TargetDeclaration.QualifiedName);
                case TargetDeclarationUsedInMultipleMethodsException usedInMultiple:
                    Logger.Warn(usedInMultiple);
                    return string.Format(RefactoringsUI.MoveCloserToUsageFailure_TargetIsUsedInMultipleMethods,
                        usedInMultiple.TargetDeclaration.QualifiedName);
                case TargetDeclarationNotUserDefinedException notUsed:
                    Logger.Warn(notUsed);
                    return string.Format(RefactoringsUI.MoveCloserToUsageFailure_TargetHasNoReferences,
                        notUsed.TargetDeclaration.QualifiedName);
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(RefactoringsUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Variable);
                default:
                    return base.Message(exception);
            }
        }
    }
}
