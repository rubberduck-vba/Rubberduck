using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeclareFieldsAsUDTMembers;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.CodeBlockInsert;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using System;

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
                    return new EncapsulateFieldRefactoringActionsProvider(
                        ResolveImpl<ReplaceReferencesRefactoringAction>(),
                        ResolveImpl<ReplacePrivateUDTMemberReferencesRefactoringAction>(),
                        ResolveImpl<ReplaceDeclarationIdentifierRefactoringAction>(),
                        ResolveImpl<DeclareFieldsAsUDTMembersRefactoringAction>(),
                        ResolveImpl<EncapsulateFieldInsertNewCodeRefactoringAction >(),
                        ResolveImpl<CodeBlockInsertRefactoringAction>()) as T;

                case nameof(EncapsulateFieldUseBackingFieldRefactoringAction):
                    return new EncapsulateFieldUseBackingFieldRefactoringAction(
                        ResolveImpl<IEncapsulateFieldRefactoringActionsProvider>(),
                        ResolveImpl<IReplacePrivateUDTMemberReferencesModelFactory>(),
                        _declarationFinderProvider, 
                        _rewritingManager) as T;

                case nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction):
                    return new EncapsulateFieldUseBackingUDTMemberRefactoringAction(
                        ResolveImpl<IEncapsulateFieldRefactoringActionsProvider>(),
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

                case nameof(IEncapsulateFieldModelFactory):
                    return new EncapsulateFieldModelFactory(
                        ResolveImpl<IEncapsulateFieldUseBackingUDTMemberModelFactory>(),
                        ResolveImpl<IEncapsulateFieldUseBackingFieldModelFactory>(),
                        ResolveImpl<IEncapsulateFieldCandidateCollectionFactory>(),
                        new EncapsulateFieldRequestFactory() as IEncapsulateFieldRequestFactory
                        ) as T;

                case nameof(IEncapsulateFieldUseBackingUDTMemberModelFactory):
                    return new EncapsulateFieldUseBackingUDTMemberModelFactory(_declarationFinderProvider,
                        ResolveImpl<IEncapsulateFieldCandidateCollectionFactory>(),
                        ResolveImpl<IObjectStateUserDefinedTypeFactory>(),
                        ResolveImpl< IEncapsulateFieldConflictFinderFactory>()) as T;

                case nameof(IEncapsulateFieldUseBackingFieldModelFactory):
                    return new EncapsulateFieldUseBackingFieldModelFactory(_declarationFinderProvider,
                        ResolveImpl<IEncapsulateFieldCandidateCollectionFactory>(),
                        ResolveImpl< IEncapsulateFieldConflictFinderFactory>()) as T;

                case nameof(IEncapsulateFieldCandidateFactory):
                    return new EncapsulateFieldCandidateFactory(_declarationFinderProvider, new CodeBuilder()) as T;

                case nameof(IObjectStateUserDefinedTypeFactory):
                    return new ObjectStateUserDefinedTypeFactory() as T;

                case nameof(IEncapsulateFieldCandidateCollectionFactory):
                    return new EncapsulateFieldCandidateCollectionFactory(_declarationFinderProvider, ResolveImpl<IEncapsulateFieldCandidateFactory>()) as T;

                case nameof(IEncapsulateFieldConflictFinderFactory):
                    return new EncapsulateFieldConflictFinderFactory(_declarationFinderProvider) as T;

            }
            throw new ArgumentException($"Unable to resolve {typeof(T).Name}") ;
        }
    }
}
