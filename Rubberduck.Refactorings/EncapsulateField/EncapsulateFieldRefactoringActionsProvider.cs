using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using Rubberduck.Refactorings.ModifyUserDefinedType;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringActionsProvider
    {
        ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers { get; }
        ICodeOnlyRefactoringAction<ModifyUserDefinedTypeModel> ModifyUserDefinedType { get; }
        ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode { get; }
    }

    /// <summary>
    /// EncapsulateFieldRefactoringActionsProvider reduces the number of EncapsulateField refactoring action 
    /// constructor parameters.  It provides Refactoring Actions common to the EncapsulateFieldRefactoringActions
    /// </summary>
    public class EncapsulateFieldRefactoringActionsProvider : IEncapsulateFieldRefactoringActionsProvider
    {
        private readonly ReplaceDeclarationIdentifierRefactoringAction _replaceDeclarationIdentifiers;
        private readonly ModifyUserDefinedTypeRefactoringAction _modifyUDTRefactoringAction;
        private readonly EncapsulateFieldInsertNewCodeRefactoringAction _encapsulateFieldInsertNewCodeRefactoringAction;

        public EncapsulateFieldRefactoringActionsProvider(
            ReplaceDeclarationIdentifierRefactoringAction replaceDeclarationIdentifierRefactoringAction,
            ModifyUserDefinedTypeRefactoringAction modifyUserDefinedTypeRefactoringAction,
            EncapsulateFieldInsertNewCodeRefactoringAction encapsulateFieldInsertNewCodeRefactoringAction)
        {
            _replaceDeclarationIdentifiers = replaceDeclarationIdentifierRefactoringAction;
            _modifyUDTRefactoringAction = modifyUserDefinedTypeRefactoringAction;
            _encapsulateFieldInsertNewCodeRefactoringAction = encapsulateFieldInsertNewCodeRefactoringAction;
        }

        public ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers 
            => _replaceDeclarationIdentifiers;

        public ICodeOnlyRefactoringAction<ModifyUserDefinedTypeModel> ModifyUserDefinedType
            => _modifyUDTRefactoringAction;

        public ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode
            => _encapsulateFieldInsertNewCodeRefactoringAction;
    }
}
