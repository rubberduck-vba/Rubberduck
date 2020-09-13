using Rubberduck.Refactorings.CreateUDTMember;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringActionsProvider
    {
        ICodeOnlyRefactoringAction<ReplaceReferencesModel> ReplaceReferences { get; }
        ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> ReplaceUDTMemberReferences { get; }
        ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers { get; }
        ICodeOnlyRefactoringAction<CreateUDTMemberModel> CreateUDTMember { get; }
        ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode { get; }
    }

    public class EncapsulateFieldRefactoringActionsProvider : IEncapsulateFieldRefactoringActionsProvider
    {
        private readonly ReplaceReferencesRefactoringAction _replaceReferences;
        private readonly ReplaceDeclarationIdentifierRefactoringAction _replaceDeclarationIdentifiers;
        private readonly ReplacePrivateUDTMemberReferencesRefactoringAction _replaceUDTMemberReferencesRefactoringAction;
        private readonly CreateUDTMemberRefactoringAction _createUDTMemberRefactoringAction;
        private readonly EncapsulateFieldInsertNewCodeRefactoringAction _encapsulateFieldInsertNewCodeRefactoringAction;

        public EncapsulateFieldRefactoringActionsProvider(
            ReplaceReferencesRefactoringAction replaceReferencesRefactoringAction,
            ReplacePrivateUDTMemberReferencesRefactoringAction replaceUDTMemberReferencesRefactoringAction,
            ReplaceDeclarationIdentifierRefactoringAction replaceDeclarationIdentifierRefactoringAction,
            CreateUDTMemberRefactoringAction createUDTMemberRefactoringActionRefactoringAction,
            EncapsulateFieldInsertNewCodeRefactoringAction encapsulateFieldInsertNewCodeRefactoringAction)
        {
            _replaceReferences = replaceReferencesRefactoringAction;
            _replaceUDTMemberReferencesRefactoringAction = replaceUDTMemberReferencesRefactoringAction;
            _replaceDeclarationIdentifiers = replaceDeclarationIdentifierRefactoringAction;
            _createUDTMemberRefactoringAction = createUDTMemberRefactoringActionRefactoringAction;
            _encapsulateFieldInsertNewCodeRefactoringAction = encapsulateFieldInsertNewCodeRefactoringAction;
        }

        public ICodeOnlyRefactoringAction<ReplaceReferencesModel> ReplaceReferences 
            => _replaceReferences;

        public ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers 
            => _replaceDeclarationIdentifiers;

        public ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> ReplaceUDTMemberReferences
            => _replaceUDTMemberReferencesRefactoringAction;

        public ICodeOnlyRefactoringAction<CreateUDTMemberModel> CreateUDTMember
            => _createUDTMemberRefactoringAction;

        public ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode
            => _encapsulateFieldInsertNewCodeRefactoringAction;
    }
}
