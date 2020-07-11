using System;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactoringAction : IRefactoringAction<RenameModel>
    {
        private readonly IRefactoringAction<RenameModel> _renameComponentOrProjectRefactoringAction;
        private readonly IRefactoringAction<RenameModel> _renameCodeDefinedIdentifierRefactoringAction;

        public RenameRefactoringAction(
            RenameComponentOrProjectRefactoringAction renameComponentOrProjectRefactoringAction,
            RenameCodeDefinedIdentifierRefactoringAction renameCodeDefinedIdentifierRefactoringAction)
        {
            _renameCodeDefinedIdentifierRefactoringAction = renameCodeDefinedIdentifierRefactoringAction;
            _renameComponentOrProjectRefactoringAction = renameComponentOrProjectRefactoringAction;
        }

        public void Refactor(RenameModel model)
        {
            if (model?.Target == null
                || model.Target.IdentifierName.Equals(model.NewName, StringComparison.InvariantCultureIgnoreCase))
            {
                return;
            }

            if (IsComponentOrProjectTarget(model))
            {
                _renameComponentOrProjectRefactoringAction.Refactor(model);
            }
            else
            {
                _renameCodeDefinedIdentifierRefactoringAction.Refactor(model);
            }
        }

        private static bool IsComponentOrProjectTarget(RenameModel model)
        { 
            var targetType = model.Target.DeclarationType;
            return targetType.HasFlag(DeclarationType.Module)
                   || targetType.HasFlag(DeclarationType.Project);
        }
    }
}