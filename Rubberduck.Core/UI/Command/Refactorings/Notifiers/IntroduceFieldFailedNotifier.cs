﻿using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.IntroduceField;
using Refactoring = Rubberduck.Resources.Refactorings;

namespace Rubberduck.UI.Command.Refactorings.Notifiers
{
    public class IntroduceFieldFailedNotifier : RefactoringFailureNotifierBase
    {
        public IntroduceFieldFailedNotifier(IMessageBox messageBox)
            : base(messageBox)
        { }

        protected override string Caption => Refactoring.Refactorings.IntroduceField_Caption;

        protected override string Message(RefactoringException exception)
        {
            switch (exception)
            {
                case TargetIsAlreadyAFieldException isAlreadyAField:
                    Logger.Warn(isAlreadyAField);
                    return string.Format(Refactoring.Refactorings.IntroduceFieldFailed_TargetIsAlreadyAField,
                        isAlreadyAField.TargetDeclaration.QualifiedName);
                case InvalidDeclarationTypeException invalidDeclarationType:
                    Logger.Warn(invalidDeclarationType);
                    return string.Format(Resources.RubberduckUI.RefactoringFailure_InvalidDeclarationType,
                        invalidDeclarationType.TargetDeclaration.QualifiedModuleName,
                        invalidDeclarationType.TargetDeclaration.DeclarationType,
                        DeclarationType.Variable);
                default:
                    return base.Message(exception);
            }
        }
    }
}