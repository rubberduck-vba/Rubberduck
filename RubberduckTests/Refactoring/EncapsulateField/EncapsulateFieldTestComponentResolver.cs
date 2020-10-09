using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using System;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;
using Rubberduck.Refactorings.ModifyUserDefinedType;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class EncapsulateFieldTestComponentResolver
    {
        private static IDeclarationFinderProvider _declarationFinderProvider;
        private static IRewritingManager _rewritingManager;
        private static QualifiedModuleName? _qmn;
        public EncapsulateFieldTestComponentResolver(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager, QualifiedModuleName? qmn = null)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _qmn = qmn;
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

                case nameof(EncapsulateFieldInsertNewCodeRefactoringAction):
                    return new EncapsulateFieldInsertNewCodeRefactoringAction(
                        _declarationFinderProvider,
                        _rewritingManager,
                        new PropertyAttributeSetsGenerator(),
                        ResolveImpl<IEncapsulateFieldCodeBuilder>()) as T;

                case nameof(ReplacePrivateUDTMemberReferencesRefactoringAction):
                    return new ReplacePrivateUDTMemberReferencesRefactoringAction(_rewritingManager) as T;

                case nameof(IEncapsulateFieldRefactoringActionsProvider):
                    return new EncapsulateFieldRefactoringActionsProvider(
                        ResolveImpl<ReplaceReferencesRefactoringAction>(),
                        ResolveImpl<ReplacePrivateUDTMemberReferencesRefactoringAction>(),
                        ResolveImpl<ReplaceDeclarationIdentifierRefactoringAction>(),
                        ResolveImpl<ModifyUserDefinedTypeRefactoringAction>(),
                        ResolveImpl<EncapsulateFieldInsertNewCodeRefactoringAction>()
                        ) as T;

                case nameof(EncapsulateFieldUseBackingFieldRefactoringAction):
                    return new EncapsulateFieldUseBackingFieldRefactoringAction(
                        ResolveImpl<IEncapsulateFieldRefactoringActionsProvider>(),
                        ResolveImpl<IReplacePrivateUDTMemberReferencesModelFactory>(),
                        _rewritingManager,
                        ResolveImpl<INewContentAggregatorFactory>()) as T;

                case nameof(EncapsulateFieldUseBackingUDTMemberRefactoringAction):
                    return new EncapsulateFieldUseBackingUDTMemberRefactoringAction(
                        ResolveImpl<IEncapsulateFieldRefactoringActionsProvider>(),
                        ResolveImpl<IReplacePrivateUDTMemberReferencesModelFactory>(),
                        _rewritingManager,
                        ResolveImpl<INewContentAggregatorFactory>()) as T;

                case nameof(IReplacePrivateUDTMemberReferencesModelFactory):
                    return new ReplacePrivateUDTMemberReferencesModelFactory(_declarationFinderProvider) as T;

                case nameof(ModifyUserDefinedTypeRefactoringAction):
                    return new ModifyUserDefinedTypeRefactoringAction(
                        _declarationFinderProvider,
                        _rewritingManager,
                        ResolveImpl<ICodeBuilder>()) as T;

                case nameof(EncapsulateFieldPreviewProvider):
                    return new EncapsulateFieldPreviewProvider(
                        ResolveImpl<EncapsulateFieldUseBackingFieldPreviewProvider>(),
                        ResolveImpl<EncapsulateFieldUseBackingUDTMemberPreviewProvider>()) as T;

                case nameof(EncapsulateFieldUseBackingFieldPreviewProvider):
                    return new EncapsulateFieldUseBackingFieldPreviewProvider(
                        ResolveImpl<EncapsulateFieldUseBackingFieldRefactoringAction>(),
                        _rewritingManager,
                        ResolveImpl<INewContentAggregatorFactory>()) as T;

                case nameof(EncapsulateFieldUseBackingUDTMemberPreviewProvider):
                    return new EncapsulateFieldUseBackingUDTMemberPreviewProvider(
                        ResolveImpl<EncapsulateFieldUseBackingUDTMemberRefactoringAction>(),
                        _rewritingManager,
                        ResolveImpl<INewContentAggregatorFactory>()) as T;

                case nameof(IEncapsulateFieldModelFactory):
                    return new EncapsulateFieldModelFactory(_declarationFinderProvider,
                        ResolveImpl<IEncapsulateFieldCandidateFactory>(),
                        ResolveImpl<IEncapsulateFieldUseBackingUDTMemberModelFactory>(),
                        ResolveImpl<IEncapsulateFieldUseBackingFieldModelFactory>(),
                        ResolveImpl<IEncapsulateFieldCandidateSetsProviderFactory>(),
                        ResolveImpl< IEncapsulateFieldConflictFinderFactory>()
                        ) as T;

                case nameof(IEncapsulateFieldUseBackingUDTMemberModelFactory):
                    return new EncapsulateFieldUseBackingUDTMemberModelFactory(_declarationFinderProvider,
                        ResolveImpl<IEncapsulateFieldCandidateFactory>(),
                        ResolveImpl<IEncapsulateFieldCandidateSetsProviderFactory>(),
                        ResolveImpl< IEncapsulateFieldConflictFinderFactory>()) as T;

                case nameof(IEncapsulateFieldUseBackingFieldModelFactory):
                    return new EncapsulateFieldUseBackingFieldModelFactory(_declarationFinderProvider,
                        ResolveImpl<IEncapsulateFieldCandidateFactory>(),
                        ResolveImpl<IEncapsulateFieldCandidateSetsProviderFactory>(),
                        ResolveImpl< IEncapsulateFieldConflictFinderFactory>()) as T;

                case nameof(IEncapsulateFieldCandidateSetsProvider):
                    if (!_qmn.HasValue)
                    {
                        throw new ArgumentException($"QualifiedModuleName is not set");
                    }
                    return new EncapsulateFieldCandidateSetsProvider(
                        _declarationFinderProvider,
                        ResolveImpl<IEncapsulateFieldCandidateFactory>(),
                        _qmn.Value) as T;

                case nameof(IEncapsulateFieldCandidateFactory):
                    return new EncapsulateFieldCandidateFactory(_declarationFinderProvider) as T;

                case nameof(IEncapsulateFieldCandidateSetsProviderFactory):
                    return new TestEFCandidateSetsProviderFactory() as T;

                case nameof(IEncapsulateFieldConflictFinderFactory):
                    return new TestEFConflictFinderFactory() as T;

                case nameof(INewContentAggregatorFactory):
                    return new TestNewContentAggregatorFactory() as T;

                case nameof(IEncapsulateFieldCodeBuilder):
                    return new EncapsulateFieldCodeBuilder(ResolveImpl<ICodeBuilder>()) as T;

                case nameof(ICodeBuilder):
                    return new CodeBuilder(ResolveImpl<IIndenter>()) as T;

                case nameof(IIndenter):
                    return new Indenter(null, CreateIndenterSettings) as T;
            }
            throw new ArgumentException($"Unable to resolve {typeof(T).Name}") ;
        }

        private static IndenterSettings CreateIndenterSettings()
        {
            var s = IndenterSettingsTests.GetMockIndenterSettings();
            s.VerticallySpaceProcedures = true;
            s.LinesBetweenProcedures = 1;
            return s;
        }
    }

    internal class TestEFCandidateSetsProviderFactory : IEncapsulateFieldCandidateSetsProviderFactory
    {
        public IEncapsulateFieldCandidateSetsProvider Create(IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldCandidateFactory factory, QualifiedModuleName qmn)
        {
            return new EncapsulateFieldCandidateSetsProvider(declarationFinderProvider, factory, qmn);
        }
    }

    internal class TestEFConflictFinderFactory : IEncapsulateFieldConflictFinderFactory
    {
        public IEncapsulateFieldConflictFinder Create(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<IEncapsulateFieldCandidate> candidates, IEnumerable<IObjectStateUDT> objectStateUDTs)
        {
            return new EncapsulateFieldConflictFinder(declarationFinderProvider, candidates, objectStateUDTs);
        }
    }

    internal class TestNewContentAggregatorFactory : INewContentAggregatorFactory
    {
        public INewContentAggregator Create()
        {
            return new NewContentAggregator();
        }
    }
}
