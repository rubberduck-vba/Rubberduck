using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeclareFieldsAsUDTMembers;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.CodeBlockInsert;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class EncapsulateFieldTestComponentResolver
    {
        private static IDeclarationFinderProvider _declarationFinderProvider;
        private static IRewritingManager _rewritingManager;
        public EncapsulateFieldTestComponentResolver(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
        }

        public T Resolve<T>() where T : class
        {
            return ResolveImpl<T>();
        }

        private static T ResolveImpl<T>() where T : class
        {
            switch (typeof(T).Name)
            {
                case nameof(EncapsulateFieldRefactoringAction):
                    return new EncapsulateFieldRefactoringAction(
                        ResolveImpl<EncapsulateFieldUseBackingFieldRefactoringAction>(), 
                        ResolveImpl<EncapsulateFieldUseBackingUDTMemberRefactoringAction>()) as T;
                case nameof(ReplaceReferencesRefactoringAction):
                    return new ReplaceReferencesRefactoringAction(_rewritingManager) as T;
                case nameof(ReplaceDeclarationIdentifierRefactoringAction):
                    return new ReplaceDeclarationIdentifierRefactoringAction(_rewritingManager) as T;
                case nameof(CodeBlockInsertRefactoringAction):
                    return new CodeBlockInsertRefactoringAction(_declarationFinderProvider, 
                        _rewritingManager, 
                        new CodeBuilder()) as T;
                case nameof(EncapsulateFieldInsertNewCodeRefactoringAction):
                    return new EncapsulateFieldInsertNewCodeRefactoringAction(
                        ResolveImpl<CodeBlockInsertRefactoringAction>(),
                        _declarationFinderProvider, 
                        _rewritingManager, 
                        new CodeBuilder()) as T;
                case nameof(ReplacePrivateUDTMemberReferencesRefactoringAction):
                    return new ReplacePrivateUDTMemberReferencesRefactoringAction(_rewritingManager) as T;
                case nameof(IEncapsulateFieldRefactoringActionsProvider):
                case nameof(EncapsulateFieldRefactoringActionsProvider):
                    return new EncapsulateFieldRefactoringActionsProvider(
                        _declarationFinderProvider, 
                        _rewritingManager,
                        ResolveImpl<ReplaceReferencesRefactoringAction>(),
                        ResolveImpl<ReplacePrivateUDTMemberReferencesRefactoringAction>(),
                        ResolveImpl<ReplaceDeclarationIdentifierRefactoringAction>(),
                        ResolveImpl<DeclareFieldsAsUDTMembersRefactoringAction>(),
                        ResolveImpl<EncapsulateFieldInsertNewCodeRefactoringAction >(),
                        ResolveImpl<CodeBlockInsertRefactoringAction>()) as T;
                case nameof(EncapsulateFieldUseBackingFieldRefactoringAction):
                    return new EncapsulateFieldUseBackingFieldRefactoringAction(
                        ResolveImpl<EncapsulateFieldRefactoringActionsProvider>(),
                        ResolveImpl<IReplacePrivateUDTMemberReferencesModelFactory>(),
                        _declarationFinderProvider, 
                        _rewritingManager) as T;
                case nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction):
                    return new EncapsulateFieldUseBackingUDTMemberRefactoringAction(
                        ResolveImpl<EncapsulateFieldRefactoringActionsProvider>(),
                        ResolveImpl<IReplacePrivateUDTMemberReferencesModelFactory>(),
                        _declarationFinderProvider, 
                        _rewritingManager,
                        new CodeBuilder()) as T;
                case nameof(IReplacePrivateUDTMemberReferencesModelFactory):
                    return new ReplacePrivateUDTMemberReferencesModelFactory(_declarationFinderProvider) as T;
                case nameof(DeclareFieldsAsUDTMembersRefactoringAction):
                    return new DeclareFieldsAsUDTMembersRefactoringAction(
                        _declarationFinderProvider, 
                        _rewritingManager, 
                        new CodeBuilder()) as T;
                case nameof(EncapsulateFieldPreviewProvider):
                    return new EncapsulateFieldPreviewProvider(
                        ResolveImpl<EncapsulateFieldUseBackingFieldPreviewProvider>(),
                        ResolveImpl<EncapsulateFieldUseBackingUDTMemberPreviewProvider>()) as T;
                case nameof(EncapsulateFieldUseBackingFieldPreviewProvider):
                    return new EncapsulateFieldUseBackingFieldPreviewProvider(
                        ResolveImpl<EncapsulateFieldUseBackingFieldRefactoringAction>(),
                        _rewritingManager) as T;
                case nameof(EncapsulateFieldUseBackingUDTMemberPreviewProvider):
                    return new EncapsulateFieldUseBackingUDTMemberPreviewProvider(
                        ResolveImpl<EncapsulateFieldUseBackingUDTMemberRefactoringAction>(), 
                        _rewritingManager) as T;
            }
            return null;
        }

        private static IIndenter CreateIndenter(IVBE vbe = null)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
    }
}
