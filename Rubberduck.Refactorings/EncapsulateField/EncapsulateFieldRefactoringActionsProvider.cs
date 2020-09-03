using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeclareFieldsAsUDTMembers;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.CodeBlockInsert;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringActionsProvider
    {
        ICodeOnlyRefactoringAction<ReplaceReferencesModel> ReplaceReferences { get; }
        ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> ReplaceUDTMemberReferences { get; }
        ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers { get; }
        ICodeOnlyRefactoringAction<CodeBlockInsertModel> CodeBlockInsert { get; }
        ICodeOnlyRefactoringAction<DeclareFieldsAsUDTMembersModel> DeclareFieldsAsUDTMembers { get; }
        ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode { get; }
    }

    public class EncapsulateFieldRefactoringActionsProvider : IEncapsulateFieldRefactoringActionsProvider
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly ReplaceReferencesRefactoringAction _replaceReferences;
        private readonly ReplaceDeclarationIdentifierRefactoringAction _replaceDeclarationIdentifiers;
        private readonly CodeBlockInsertRefactoringAction _codeBlockInsertRefactoringAction;
        private readonly ReplacePrivateUDTMemberReferencesRefactoringAction _replaceUDTMemberReferencesRefactoringAction;
        private readonly DeclareFieldsAsUDTMembersRefactoringAction _declareFieldsAsUDTMembersRefactoringAction;
        private readonly EncapsulateFieldInsertNewCodeRefactoringAction _encapsulateFieldInsertNewCodeRefactoringAction;

        public EncapsulateFieldRefactoringActionsProvider(IDeclarationFinderProvider declarationFinderProvider, 
            IRewritingManager rewritingManager,
            ReplaceReferencesRefactoringAction replaceReferencesRefactoringAction,
            ReplacePrivateUDTMemberReferencesRefactoringAction replaceUDTMemberReferencesRefactoringAction,
            ReplaceDeclarationIdentifierRefactoringAction replaceDeclarationIdentifierRefactoringAction,
            DeclareFieldsAsUDTMembersRefactoringAction declareFieldsAsUDTMembersRefactoringAction,
            EncapsulateFieldInsertNewCodeRefactoringAction encapsulateFieldInsertNewCodeRefactoringAction,
            CodeBlockInsertRefactoringAction codeBlockInsertRefactoringAction)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _replaceReferences = replaceReferencesRefactoringAction;
            _replaceUDTMemberReferencesRefactoringAction = replaceUDTMemberReferencesRefactoringAction;
            _replaceDeclarationIdentifiers = replaceDeclarationIdentifierRefactoringAction;
            _declareFieldsAsUDTMembersRefactoringAction = declareFieldsAsUDTMembersRefactoringAction;
            _codeBlockInsertRefactoringAction = codeBlockInsertRefactoringAction;
            _encapsulateFieldInsertNewCodeRefactoringAction = encapsulateFieldInsertNewCodeRefactoringAction;
        }

        public ICodeOnlyRefactoringAction<ReplaceReferencesModel> ReplaceReferences 
            => _replaceReferences;

        public ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> ReplaceDeclarationIdentifiers 
            => _replaceDeclarationIdentifiers;

        public ICodeOnlyRefactoringAction<CodeBlockInsertModel> CodeBlockInsert
            => _codeBlockInsertRefactoringAction;

        public ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> ReplaceUDTMemberReferences
            => _replaceUDTMemberReferencesRefactoringAction;

        public ICodeOnlyRefactoringAction<DeclareFieldsAsUDTMembersModel> DeclareFieldsAsUDTMembers
            => _declareFieldsAsUDTMembersRefactoringAction;

        public ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> EncapsulateFieldInsertNewCode
            => _encapsulateFieldInsertNewCodeRefactoringAction;
    }
}
