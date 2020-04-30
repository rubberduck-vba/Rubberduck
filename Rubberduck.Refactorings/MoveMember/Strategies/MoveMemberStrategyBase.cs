using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public abstract class MoveMemberStrategyBase
    {
        protected enum RequiredGroup
        {
            PrivateRetain,
            PrivateMove,
            PublicRetain,
            PublicMove,
        }

        protected enum MoveDisposition
        {
            Move,
            Retain
        }

        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected readonly RenameCodeDefinedIdentifierRefactoringAction _renameAction;
        protected readonly IMoveMemberMoveGroupsProviderFactory _moveGroupsProviderFactory;

        protected readonly IConflictDetectionSessionFactory _conflictDetectionSessionFactory;
        protected readonly IConflictDetectionDeclarationProxyFactory _declarationProxyFactory;

        protected MoveMemberStrategyBase(IDeclarationFinderProvider declarationFinderProvider,
                                        RenameCodeDefinedIdentifierRefactoringAction renameAction,
                                        IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory,
                                        IConflictDetectionSessionFactory namingToolsSessionFactory,
                                        IConflictDetectionDeclarationProxyFactory declarationProxyFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _renameAction = renameAction;
            _moveGroupsProviderFactory = moveGroupsProviderFactory;
            _declarationProxyFactory = declarationProxyFactory;
            _conflictDetectionSessionFactory = namingToolsSessionFactory;
        }

        public abstract bool IsApplicable(MoveMemberModel model);

        protected virtual bool IsExecutableModelBase(MoveMemberModel model, out string nonExecutableMessage)
        {
            nonExecutableMessage = string.Empty;

            if (!model.SelectedDeclarations.Any())
            {
                nonExecutableMessage = Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove;
                return false;
            }

            if (string.IsNullOrEmpty(model?.Destination.ModuleName))
            {
                nonExecutableMessage = Resources.RubberduckUI.MoveMember_UndefinedDestinationModule;
                return false;
            }

            if (!IsApplicable(model))
            {
                nonExecutableMessage = Resources.RubberduckUI.MoveMember_ApplicableStrategyNotFound;
                return false;
            }
            return true;
        }

        public virtual void RefactorRewrite(MoveMemberModel model, IRewriteSession moveMemberRewriteSession, IRewritingManager rewritingManager, INewContentProvider newContentProvider, out string newModuleContent)
        {
            newModuleContent = string.Empty;

            var moveGroups = _moveGroupsProviderFactory.Create(model.MoveableMemberSets);

            var dispositions = DetermineDispositionGroups(model, moveGroups);

            var scratchPadSession = rewritingManager.CheckOutCodePaneSession();

            newContentProvider = LoadNewContentProvider(newContentProvider, model, moveGroups, moveMemberRewriteSession, scratchPadSession, dispositions);
            newModuleContent = newContentProvider.AsSingleBlock;

            if (model.Destination.IsExistingModule(out _))
            {
                InsertNewContent(moveMemberRewriteSession, model.Destination, newModuleContent);
            }

            RemoveDeclarations(moveMemberRewriteSession, dispositions[MoveDisposition.Move]);

            ModifyRetainedReferencesToMovedMembers(moveMemberRewriteSession, model, dispositions);

            newContentProvider = LoadSourceNewContentProvider(newContentProvider, model);

            InsertNewContent(moveMemberRewriteSession, model.Source, newContentProvider.AsSingleBlock);

            UpdateReferencesToMovedMembersInNonEndpointModules(model, moveMemberRewriteSession, dispositions);
        }

        protected abstract Dictionary<MoveDisposition, List<Declaration>> DetermineDispositionGroups(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups);

        protected bool TrySetDispositionGroupsForStandardModuleSource(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, out Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            dispositions = EmptyDispositions();

            //Any Private declaration (member, field, or constant) that is used by both the Move Participants
            //and NonParticipants MUST be retained in the Source module.  If a Selected Element (which HAS to move)
            //directly references a declaration that MUST be retained - the move cannot be executed
            if (DependenciesRequiredToMove(moveGroups).Intersect(PrivateDeclarationsThatMustBeRetained(moveGroups)).Any())
            {
                return false;
            }

            var supportingDeclarations = new Dictionary<RequiredGroup, List<Declaration>>()
            {
                [RequiredGroup.PrivateMove] = new List<Declaration>(),
                [RequiredGroup.PrivateRetain] = PrivateDeclarationsThatMustBeRetained(moveGroups),
                [RequiredGroup.PublicMove] = new List<Declaration>(),
                [RequiredGroup.PublicRetain] = moveGroups.Declarations(MoveGroup.Support_Public).ToList(),
            };

            //MS-VBAL 5.2.3.1
            //The declared type of a public variable defined in a class module may not be a private UDT or Enum.
            if (model.Destination.IsClassModule)
            {
                var publicVariablesThatMustMove = moveGroups.Declarations(MoveGroup.Selected).Where(mm => mm.IsMemberVariable() && !mm.HasPrivateAccessibility());
                if (publicVariablesThatMustMove.Any(publicVariable => publicVariable.IsUserDefinedTypeField() || publicVariable.IsEnumField()))
                {
                    return false;
                }
            }

            var mustMoveDependencies = DependenciesRequiredToMove(moveGroups);
            foreach (var selectedMoveableMemberSetDependency in moveGroups.ToMoveableMemberSets(mustMoveDependencies))
            {
                //All directly referenced Private declarations of a Selected element must be moved.
                //Otherwise, the Selected element cannot be moved
                if (selectedMoveableMemberSetDependency.HasPrivateAccessibility)
                {
                    supportingDeclarations[RequiredGroup.PrivateMove].AddRange(selectedMoveableMemberSetDependency.Members);
                }
            }

            var publicSupportMoveables = moveGroups.MoveableMemberSets(MoveGroup.Support_Public);
            for (var idx = 0; idx < publicSupportMoveables.Count; idx++)
            {
                var moveable = publicSupportMoveables.ElementAt(idx);
                if (!supportingDeclarations[RequiredGroup.PublicRetain].Contains(moveable.Member))
                {
                    continue;
                }

                //If a Public support members references a Private support declaration that 'must move'
                //then the Public support member 'must move' as well.
                if (IsMustMovePublicSupport(moveable, supportingDeclarations[RequiredGroup.PrivateMove], out var newPrivateMustMoveSupport))
                {
                    supportingDeclarations[RequiredGroup.PublicMove].AddRange(moveable.Members);

                    supportingDeclarations[RequiredGroup.PublicRetain] = supportingDeclarations[RequiredGroup.PublicRetain]
                                                                            .Except(moveable.Members)
                                                                            .ToList();

                    supportingDeclarations[RequiredGroup.PrivateMove] = supportingDeclarations[RequiredGroup.PrivateMove]
                                                                            .Concat(newPrivateMustMoveSupport)
                                                                            .Distinct()
                                                                            .ToList();

                    //Need to work from the start of the list again to see if the added private support
                    //dependencies forces a move of any other MoveableMembers in the 'Retain' collection 
                    idx = -1;
                }
            }

            var retainedPrivateSupportCandidates = moveGroups.Declarations(MoveGroup.Support_Private).Except(supportingDeclarations[RequiredGroup.PrivateMove]);

            var privateDependenciesOfRetainedPublicSupportMembers = retainedPrivateSupportCandidates.AllReferences()
                            .Where(rf => supportingDeclarations[RequiredGroup.PublicRetain].Contains(rf.ParentScoping))
                            .Select(rf => rf.Declaration);

            supportingDeclarations[RequiredGroup.PrivateRetain].AddRange(privateDependenciesOfRetainedPublicSupportMembers);

            var privateExclusiveSupportMoveableMemberSets = moveGroups.MoveableMemberSets(MoveGroup.Support_Exclusive)
                                .Where(p => p.HasPrivateAccessibility)
                                .Except(moveGroups.ToMoveableMemberSets(supportingDeclarations[RequiredGroup.PrivateRetain]));

            //if Private exclusive support members (which must move) have
            //a declaration in common with the Private declarations that must be retained in the Source
            //Module...the move is not executable.
            if (privateExclusiveSupportMoveableMemberSets.SelectMany(mm => mm.DirectDependencies)
                .Intersect(supportingDeclarations[RequiredGroup.PrivateRetain]).Any())
            {
                return false;
            }

            supportingDeclarations[RequiredGroup.PrivateMove].AddRange(privateExclusiveSupportMoveableMemberSets.SelectMany(mm => mm.Members));

            foreach (var key in supportingDeclarations.Keys.ToList())
            {
                supportingDeclarations[key] = supportingDeclarations[key].Distinct().ToList();
            }

            //Final check to see that all 'binning' has not resulted in overlaps.  If there are overlaps,
            //the strategy cannot fully resolve the scenario and execute the move 
            if (supportingDeclarations[RequiredGroup.PrivateMove].Intersect(supportingDeclarations[RequiredGroup.PrivateRetain]).Any()
                || supportingDeclarations[RequiredGroup.PublicMove].Intersect(supportingDeclarations[RequiredGroup.PublicRetain]).Any())
            {
                return false;
            }

            dispositions[MoveDisposition.Move] = (moveGroups.Declarations(MoveGroup.Selected)
                                                                        .Concat(supportingDeclarations[RequiredGroup.PublicMove])
                                                                        .Concat(supportingDeclarations[RequiredGroup.PrivateMove])).ToList();

            dispositions[MoveDisposition.Retain] = (moveGroups.Declarations(MoveGroup.AllParticipants)
                                                                        .Except(dispositions[MoveDisposition.Move])).ToList();
            return true;
        }

        private INewContentProvider LoadNewContentProvider(INewContentProvider contentProvider, MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, IRewriteSession moveMemberRewriteSession, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            contentProvider.ResetContent();

            if (model.Destination.IsExistingModule(out var destinationModule))
            {
                RenameDestinationNameConflicts(model, moveMemberRewriteSession, scratchPadSession, dispositions[MoveDisposition.Move]);

                ModifyExistingReferencesToMovedMembersInDestination(destinationModule, moveMemberRewriteSession, dispositions);
            }

            foreach (var element in dispositions[MoveDisposition.Move])
            {
                if (element.IsMember())
                {
                    var memberCodeBlock = CreateMovedMemberCodeBlock(model, moveGroups, element, scratchPadSession, dispositions);
                    contentProvider.AddMember(memberCodeBlock);
                    continue;
                }

                if (element.DeclarationType.Equals(DeclarationType.UserDefinedType)
                    || element.DeclarationType.Equals(DeclarationType.Enumeration))
                {
                    var rewriter = scratchPadSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);
                    contentProvider.AddTypeDeclaration(rewriter.GetText(element));
                    continue;
                }

                var nonMembercodeBlock = CreateMovedNonMemberCodeBlock(model, moveGroups, element, scratchPadSession, dispositions);
                contentProvider.AddFieldOrConstantDeclaration(nonMembercodeBlock);
            }

            return contentProvider;
        }

        protected static void InsertNewContent(IRewriteSession refactoringRewriteSession, IMoveMemberEndpoint endpoint, string movedContent)
        {
            if (endpoint is IMoveDestinationEndpoint destination)
            {
                if (!destination.IsExistingModule(out var module))
                {
                    throw new MoveMemberUnsupportedMoveException();
                }

                var destinationRewriter = refactoringRewriteSession.CheckOutModuleRewriter(module.QualifiedModuleName);

                if (endpoint.TryGetCodeSectionStartIndex(out var destCoodeSectionStartIndex))
                {
                    destinationRewriter.InsertBefore(destCoodeSectionStartIndex, $"{movedContent}{Environment.NewLine}{Environment.NewLine}");
                }
                else
                {
                    destinationRewriter.InsertAtEndOfFile($"{Environment.NewLine}{Environment.NewLine}{movedContent}");
                }
            }

            if (endpoint is IMoveSourceEndpoint sourceEndpoint)
            {
                var sourceRewriter = refactoringRewriteSession.CheckOutModuleRewriter(sourceEndpoint.QualifiedModuleName);

                if (endpoint.TryGetCodeSectionStartIndex(out var sourceCodeSectionStartIndex))
                {
                    sourceRewriter.InsertBefore(sourceCodeSectionStartIndex, $"{movedContent}{Environment.NewLine}{Environment.NewLine}");
                }
                else
                {
                    sourceRewriter.InsertAtEndOfFile($"{Environment.NewLine}{Environment.NewLine}{movedContent}");
                }
            }
        }

        protected virtual void ModifyRetainedReferencesToMovedMembers(IRewriteSession rewriteSession, MoveMemberModel model, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var renamableReferences = RenameableReferences(dispositions[MoveDisposition.Move], model.Source.QualifiedModuleName);
            var retainedReferencesToModuleQualify = renamableReferences.Where(rf => !dispositions[MoveDisposition.Move].Contains(rf.ParentScoping));

            var moveableConstants = model.MoveableMemberSets.Where(mm => mm.Member.IsConstant());
            var directReferencesOfMovedConstants = new List<IdentifierReference>();
            foreach (var constant in moveableConstants)
            {
                if (dispositions[MoveDisposition.Move].Contains(constant.Member))
                {
                    directReferencesOfMovedConstants.AddRange(constant.DirectReferences);
                    retainedReferencesToModuleQualify = retainedReferencesToModuleQualify.Except(constant.DirectReferences);
                }
            }

            var moveableFields = model.MoveableMemberSets.Where(mm => mm.Member.IsMemberVariable());
            var directReferencesOfMovedFields = new List<IdentifierReference>();
            foreach (var field in moveableFields)
            {
                if (dispositions[MoveDisposition.Move].Contains(field.Member))
                {
                    directReferencesOfMovedFields.AddRange(field.DirectReferences);
                    retainedReferencesToModuleQualify = retainedReferencesToModuleQualify.Except(field.DirectReferences);
                }
            }

            if (retainedReferencesToModuleQualify.Any())
            {
                var sourceRewriter = rewriteSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);
                foreach (var rf in retainedReferencesToModuleQualify)
                {
                    sourceRewriter.Replace(rf.Context, AddDestinationModuleQualification(model, rf, dispositions[MoveDisposition.Retain]));
                }
            }
        }

        protected abstract INewContentProvider LoadSourceNewContentProvider(INewContentProvider contentProvider, MoveMemberModel model);

        protected static IEnumerable<IdentifierReference> RenameableReferences(IEnumerable<Declaration> declarations, QualifiedModuleName qmn)
                    => RenameableReferencesByQualifiedModuleName(declarations.AllReferences())
                                                    .Where(g => qmn == g.Key)
                                                    .SelectMany(g => g);

        protected static Dictionary<MoveDisposition, List<Declaration>> EmptyDispositions()
        {
            return new Dictionary<MoveDisposition, List<Declaration>>()
            {
                [MoveDisposition.Move] = new List<Declaration>(),
                [MoveDisposition.Retain] = new List<Declaration>()
            };
        }

        protected string AddDestinationModuleQualification(MoveMemberModel model, IdentifierReference identifierReference, IEnumerable<Declaration> retain)
        {
            var movedIdentifier = model.MoveableMemberSetByName(identifierReference.IdentifierName).MovedIdentifierName;

            if (NeverAddMemberAccessTypes.Contains(identifierReference.Declaration.DeclarationType)
                || (identifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Function)
                        && identifierReference.IsAssignment)
                || retain.Contains(identifierReference.Declaration))
            {
                return movedIdentifier;
            }

            return $"{model.Destination.ModuleName}.{movedIdentifier}";
        }

        private void RenameDestinationNameConflicts(MoveMemberModel model, IRewriteSession moveMemberRewriteSession, IRewriteSession scratchPadSession, IEnumerable<Declaration> movers)
        {
            var conflictSession = _conflictDetectionSessionFactory.Create();

            foreach (var mover in movers)
            {
                if (conflictSession.TryProposedRelocation(mover, model.Destination.ModuleName))
                {
                    var renamePairs = conflictSession.ConflictFreeRenamePairs;
                    foreach ((Declaration target, string newName) in renamePairs)
                    {
                        var renameModel = new RenameModel(target) { NewName = newName };
                        _renameAction.Refactor(renameModel, moveMemberRewriteSession);
                        _renameAction.Refactor(renameModel, scratchPadSession);
                    }
                }
            }
        }

        private static string CreateMovedMemberCodeBlock(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, Declaration member, IRewriteSession rewriteSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            Debug.Assert(member.IsMember());

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);
            if (member is ModuleBodyElementDeclaration mbed)
            {
                var argListContext = member.Context.GetDescendent<VBAParser.ArgListContext>();
                rewriter.Replace(argListContext, $"({mbed.ImprovedArgumentList()})");
            }

            if (moveGroups.Declarations(MoveGroup.Selected).Contains(member))
            {
                var accessibility = IsOnlyReferencedByMovedMethods(member, dispositions[MoveDisposition.Move])
                    ? member.Accessibility == Accessibility.Implicit ? Tokens.Public : member.Accessibility.TokenString()
                    : Tokens.Public;

                rewriter.SetMemberAccessibility(member, accessibility);
            }

            var otherMoveParticipantReferencesRelatedToMember = moveGroups.Declarations(MoveGroup.Support_Exclusive)
                                    .Where(esd => !esd.IsMember()).AllReferences()
                                    .Where(rf => rf.ParentScoping == member);

            if (model.Source.IsStandardModule)
            {
                AddSourceModuleQualificationToMovedReferences(member, model.Source.ModuleName, rewriter, dispositions);
            }

            var destinationMemberAccessReferencesToMovedMembers = model.Destination.ModuleDeclarations
                .AllReferences().Where(rf => rf.ParentScoping == member);

            rewriter.RemoveMemberAccess(destinationMemberAccessReferencesToMovedMembers);

            rewriter.RemoveWithMemberAccess(destinationMemberAccessReferencesToMovedMembers);

            return rewriter.GetText(member);
        }

        private static bool IsOnlyReferencedByMovedMethods(Declaration element, IEnumerable<Declaration> move)
            => element.References.All(rf => move.Where(m => m.IsMember()).Contains(rf.ParentScoping));

        private static void AddSourceModuleQualificationToMovedReferences(Declaration member, string sourceModuleName, IModuleRewriter scratchPadRewriter, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var retainedPublicDeclarations = dispositions[MoveDisposition.Retain].Where(m => !m.HasPrivateAccessibility());
            if (retainedPublicDeclarations.Any())
            {
                var destinationRefs = retainedPublicDeclarations.AllReferences().Where(rf => dispositions[MoveDisposition.Move].Contains(rf.ParentScoping));
                foreach (var rf in destinationRefs)
                {
                    scratchPadRewriter.Replace(rf.Context, $"{sourceModuleName}.{rf.IdentifierName}");
                }
            }
        }

        private static IEnumerable<IdentifierReference> ReferencesInConstantDeclarationExpressions(IMoveMemberGroupsProvider moveGroups, Declaration declaration)
        {
            var references = new List<IdentifierReference>();

            if (!declaration.IsConstant()) { return Enumerable.Empty<IdentifierReference>(); }

            var allModuleConstants = moveGroups.Declarations(MoveGroup.AllParticipants).Concat(moveGroups.Declarations(MoveGroup.NonParticipants))
                .Where(d => d.IsConstant() && d != declaration);

            foreach (var constant in allModuleConstants)
            {
                var lExprContexts = constant.Context.GetDescendents<VBAParser.LExprContext>();
                if (lExprContexts.Any())
                {
                    references.AddRange(declaration.References.Where(rf => lExprContexts.Contains(rf.Context.Parent)));
                }
            }
            return references;
        }

        private static void UpdateReferencesToMovedMembersInNonEndpointModules(MoveMemberModel model, IRewriteSession rewriteSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var endpointQMNs = new List<QualifiedModuleName>() { model.Source.QualifiedModuleName };
            if (model.Destination.IsExistingModule(out var destination))
            {
                endpointQMNs.Add(destination.QualifiedModuleName);
            }

            var qmnToReferenceGroups
                    = RenameableReferencesByQualifiedModuleName(dispositions[MoveDisposition.Move].AllReferences())
                            .Where(qmn => !endpointQMNs.Contains(qmn.Key));

            foreach (var referenceGroup in qmnToReferenceGroups)
            {
                var moduleRewriter = rewriteSession.CheckOutModuleRewriter(referenceGroup.Key);

                var idRefMemberAccessExpressionContextPairs = referenceGroup.Where(rf => rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)
                        .Select(rf => (rf, rf.Context.Parent as VBAParser.MemberAccessExprContext));

                var destinationModuleName = model.Destination.ModuleName;
                foreach (var (IdRef, MemberAccessExpressionContext) in idRefMemberAccessExpressionContextPairs)
                {
                    moduleRewriter.Replace(MemberAccessExpressionContext.lExpression(), destinationModuleName);
                }

                var idRefWithMemberAccessExprContextPairs = referenceGroup.Where(rf => rf.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        .Select(rf => (rf, rf.Context.Parent as VBAParser.WithMemberAccessExprContext));

                foreach (var (IdRef, withMemberAccessExprContext) in idRefWithMemberAccessExprContextPairs)
                {
                    moduleRewriter.InsertBefore(withMemberAccessExprContext.Start.TokenIndex, destinationModuleName);
                }

                var nonQualifiedReferences = referenceGroup.Where(rf => !(rf.Context.Parent is VBAParser.WithMemberAccessExprContext
                    || (rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)));

                foreach (var rf in nonQualifiedReferences)
                {
                    moduleRewriter.InsertBefore(rf.Context.Start.TokenIndex, $"{destinationModuleName}.");
                }
            }
        }

        private static string CreateMovedNonMemberCodeBlock(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, Declaration nonMember, IRewriteSession rewriteSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            Debug.Assert(!nonMember.IsMember());

            var moveableMember = model.MoveableMemberSetByName(nonMember.IdentifierName);

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);

            var visibility = nonMember.Accessibility.TokenString();

            if (moveableMember.IsSelected && nonMember.HasPrivateAccessibility())
            {
                var refsUsedByConstantDeclarations = ReferencesInConstantDeclarationExpressions(moveGroups, nonMember);

                var refsUsedByMovedMembers = nonMember.References.Where(rf => dispositions[MoveDisposition.Move].Where(m => m.IsMember()).Contains(rf.ParentScoping));

                var refsOtherThanMoveParticipants = nonMember.References.Except(refsUsedByConstantDeclarations.Concat(refsUsedByMovedMembers));

                visibility = refsOtherThanMoveParticipants.Any() ? Tokens.Public : visibility;
            }

            if (model.Source.IsStandardModule && visibility.Equals(Tokens.Public))
            {
                foreach (var rf in moveableMember.DirectReferences)
                {
                    if (!dispositions[MoveDisposition.Move].Contains(rf.Declaration) && !NeverAddMemberAccessTypes.Contains(rf.Declaration.DeclarationType))
                    {
                        rewriter.InsertBefore(rf.Context.Start.TokenIndex, $"{model.Source.ModuleName}.");
                    }
                }
            }

            if (moveableMember.IsEnumeration || moveableMember.IsUserDefinedType)
            {
                return rewriter.GetText(nonMember);
            }

            return nonMember.IsVariable()
                ? $"{visibility} {rewriter.GetText(nonMember)}"
                : $"{visibility} {Tokens.Const} {rewriter.GetText(nonMember)}";
        }

        private static IEnumerable<IGrouping<QualifiedModuleName, IdentifierReference>> RenameableReferencesByQualifiedModuleName(IEnumerable<IdentifierReference> references)
        {
            //The filter used by RenameRefactoring
            var renamableReferencesByQMN = references
                .Where(reference =>
                    reference.Context.GetText() != Tokens.Me
                    && !reference.IsArrayAccess
                    && !reference.IsDefaultMemberAccess)
                .GroupBy(r => r.QualifiedModuleName);

            return renamableReferencesByQMN;
        }

        /// <summary>
        /// Clears entire VariableStmtContext or ConstantStmtContext
        /// when all the variables or constants declared in the list are removed.
        /// </summary>
        /// <param name="rewriteSession"></param>
        /// <param name="declarations"></param>
        private static void RemoveDeclarations(IRewriteSession rewriteSession, IEnumerable<Declaration> declarations)
        {
            var removedVariables = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();
            var removedConstants = new Dictionary<VBAParser.ConstStmtContext, HashSet<Declaration>>();

            foreach (var declaration in declarations)
            {
                if (declaration.DeclarationType.Equals(DeclarationType.Variable))
                {
                    CacheListDeclaredElement<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(rewriteSession, declaration, removedVariables);
                    continue;
                }

                if (declaration.DeclarationType.Equals(DeclarationType.Constant))
                {
                    CacheListDeclaredElement<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(rewriteSession, declaration, removedConstants);
                    continue;
                }

                var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
                rewriter.Remove(declaration);
            }

            ExecuteCachedRemoveRequests<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(rewriteSession, removedVariables);
            ExecuteCachedRemoveRequests<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(rewriteSession, removedConstants);
        }

        private static void CacheListDeclaredElement<T, K>(IRewriteSession rewriteSession, Declaration target, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            var declarationList = target.Context.GetAncestor<T>();

            if ((declarationList?.children.OfType<K>().Count() ?? 1) == 1)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
                rewriter.Remove(target);
                return;
            }

            if (!dictionary.ContainsKey(declarationList))
            {
                dictionary.Add(declarationList, new HashSet<Declaration>());
            }
            dictionary[declarationList].Add(target);
        }

        private static void ExecuteCachedRemoveRequests<T, K>(IRewriteSession rewriteSession, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            foreach (var key in dictionary.Keys.Where(k => dictionary[k].Any()))
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(dictionary[key].First().QualifiedModuleName);

                if (key.children.OfType<K>().Count() == dictionary[key].Count)
                {
                    rewriter.Remove(key.Parent);
                    continue;
                }

                foreach (var dec in dictionary[key])
                {
                    rewriter.Remove(dec);
                }
            }
        }

        private static void ModifyExistingReferencesToMovedMembersInDestination(Declaration destination, IRewriteSession rewriteSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var destinationReferencesToMovedMembers = dispositions[MoveDisposition.Move].AllReferences()
                .Where(rf => rf.QualifiedModuleName == destination.QualifiedModuleName);

            if (destinationReferencesToMovedMembers.Any())
            {
                var destinationRewriter = rewriteSession.CheckOutModuleRewriter(destination.QualifiedModuleName);

                destinationRewriter.RemoveMemberAccess(destinationReferencesToMovedMembers);

                destinationRewriter.RemoveWithMemberAccess(destinationReferencesToMovedMembers);
            }
        }

        //Any Private declaration (member, field, or constant) that is used by both the Participants
        //and the NonParticipants MUST be retained in the Source module
        private static List<Declaration> PrivateDeclarationsThatMustBeRetained(IMoveMemberGroupsProvider moveGroups)
        {
            var allParticipantDependencies = moveGroups.Dependencies(MoveGroup.AllParticipants);
            var nonParticipantDirectDependencies = moveGroups.DirectDependencies(MoveGroup.NonParticipants);

            return allParticipantDependencies.Intersect(nonParticipantDirectDependencies)
                .Where(d => d.HasPrivateAccessibility()).ToList();
        }

        private static List<Declaration> DependenciesRequiredToMove(IMoveMemberGroupsProvider moveGroups)
            => moveGroups.DirectDependencies(MoveGroup.Selected).Intersect(moveGroups.Declarations(MoveGroup.Support_Private)).ToList();

        private static bool IsMustMovePublicSupport(IMoveableMemberSet publicSupportMember, IEnumerable<Declaration> mustMovePrivateSupport, out List<Declaration> newMustMovePrivateSupport)
        {
            newMustMovePrivateSupport = new List<Declaration>();
            foreach (var mustMovePvtSupport in mustMovePrivateSupport)
            {
                if (publicSupportMember.DirectDependencies.Contains(mustMovePvtSupport))
                {
                    //If the Public support member directly references a Private support declaration
                    //that 'has to' move, then we will include the Public support member in the moved declarations.
                    //But, now also 'have to' move the direct Private dependencies of the Public support member.
                    newMustMovePrivateSupport = publicSupportMember.DirectDependencies.Where(d => d.HasPrivateAccessibility()).ToList();
                    return true;
                }
            }
            return false;
        }

        private static List<DeclarationType> NeverAddMemberAccessTypes = new List<DeclarationType>()
        {
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember
        };
    }
}
