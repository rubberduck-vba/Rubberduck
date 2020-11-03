using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using Rubberduck.Refactorings.ModifyUserDefinedType;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringActionsProvider
    {
        ICodeOnlyRefactoringAction<ReplaceReferencesModel> ReplaceReferences { get; }
        ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> ReplaceUDTMemberReferences { get; }
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
        private readonly ReplaceReferencesRefactoringAction _replaceReferences;
        private readonly ReplaceDeclarationIdentifierRefactoringAction _replaceDeclarationIdentifiers;
        private readonly ReplacePrivateUDTMemberReferencesRefactoringAction _replaceUDTMemberReferencesRefactoringAction;
        private readonly ModifyUserDefinedTypeRefactoringAction _modifyUDTRefactoringAction;
        private readonly EncapsulateFieldInsertNewCodeRefactoringAction _encapsulateFieldInsertNewCodeRefactoringAction;

        public EncapsulateFieldRefactoringActionsProvider(
            ReplaceReferencesRefactoringAction replaceReferencesRefactoringAction,
            ReplacePrivateUDTMemberReferencesRefactoringAction replaceUDTMemberReferencesRefactoringAction,
            ReplaceDeclarationIdentifierRefactoringAction replaceDeclarationIdentifierRefactoringAction,
            ModifyUserDefinedTypeRefactoringAction modifyUserDefinedTypeRefactoringAction,
            EncapsulateFieldInsertNewCodeRefactoringAction encapsulateFieldInsertNewCodeRefactoringAction)
        {
            _replaceReferences = replaceReferencesRefactoringAction;
            _replaceUDTMemberReferencesRefactoringAction = replaceUDTMemberReferencesRefactoringAction;
            _replaceDeclarationIdentifiers = replaceDeclarationIdentifierRefactoringAction;
            _modifyUDTRefactoringAction = modifyUserDefinedTypeRefactoringAction;
            _encapsulateFieldInsertNewCodeRefactoringAction = encapsulateFieldInsertNewCodeRefactoringAction;
        }

        public ICodeOnlyRefactoringAction<ReplaceReferencesModel> ReplaceReferences 
            => _replaceReferences;

        public ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers 
            => _replaceDeclarationIdentifiers;

        public ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> ReplaceUDTMemberReferences
            => _replaceUDTMemberReferencesRefactoringAction;

        public ICodeOnlyRefactoringAction<ModifyUserDefinedTypeModel> ModifyUserDefinedType
            => _modifyUDTRefactoringAction;

        public ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode
            => _encapsulateFieldInsertNewCodeRefactoringAction;
    }
}
