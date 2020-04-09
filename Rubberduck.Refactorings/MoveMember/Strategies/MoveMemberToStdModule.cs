using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
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
    public class MoveMemberToStdModule : MoveMemberStrategyBase
    {
        private enum RequiredGroup
        {
            PrivateRetain,
            PrivateMove,
            PublicRetain,
            PublicMove,
        }

        private enum MoveDisposition
        {
            Move,
            Retain
        }

        private readonly RenameCodeDefinedIdentifierRefactoringAction _renameAction;
        private readonly IMoveMemberMoveGroupsProviderFactory _moveGroupsProviderFactory;
        public MoveMemberToStdModule(RenameCodeDefinedIdentifierRefactoringAction renameAction, 
                                        IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory)
        {
            _renameAction = renameAction;
            _moveGroupsProviderFactory = moveGroupsProviderFactory;
        }

        public override bool IsApplicable(MoveMemberModel model)
        {
            //Note: A StandardModule is the default
            //model.Destination.IsStandardModule returns true for non-specified destinations
            if (!model.Destination.IsStandardModule) { return false; }

            //Strategy does not support Private Fields as a SelectedDeclaration
            if (model.SelectedDeclarations.Any(sd => sd.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && sd.HasPrivateAccessibility())) { return false; }

            var moveGroups = _moveGroupsProviderFactory.Create(model.MoveableMembers);

            if (model.Source.IsStandardModule)
            {
                return TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out _);
            }

            return TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out _);
        }

        public override bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage)
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
            return true;
        }

        public override void RefactorRewrite(MoveMemberModel model, IRewriteSession moveMemberRewriteSession, IRewritingManager rewritingManager, IMovedContentProvider contentProvider, out string newModuleContent)
        {
            newModuleContent = string.Empty;

            var moveGroups = _moveGroupsProviderFactory.Create(model.MoveableMembers);

            var dispositions = DetermineDispositionGroups(model, moveGroups);

            var scratchPadSession = rewritingManager.CheckOutCodePaneSession();

            ClearMoveNameConflicts(model, _renameAction, moveMemberRewriteSession, scratchPadSession, dispositions);

            RemoveDeclarations(moveMemberRewriteSession, dispositions[MoveDisposition.Move]);

            if (model.Destination.IsExistingModule(out var destinationModule))
            {
                ModifyExistingReferencesToMovedMembersInDestination(destinationModule, moveMemberRewriteSession, dispositions);

                newModuleContent = InsertDestinationContent(model, moveGroups, moveMemberRewriteSession, scratchPadSession, dispositions, contentProvider);
            }
            else
            {
                contentProvider = LoadMovedContentProvider(contentProvider, model, moveGroups, scratchPadSession, dispositions);
                newModuleContent = contentProvider.AsSingleBlock;
            }

            ModifyRetainedReferencesToMovedMembers(moveMemberRewriteSession, model, dispositions);

            UpdateReferencesToMovedMembersInNonEndpointModules(model, moveMemberRewriteSession, dispositions);
        }

        private static Dictionary<MoveDisposition, List<Declaration>> DetermineDispositionGroups(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups)
        {
            var dispositions = EmptyDispositions();
            if (model.Source.IsStandardModule
                && TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out dispositions))
            {
                return dispositions;
            }

            if (TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out dispositions))
            {
                return dispositions;
            }

            throw new MoveMemberUnsupportedMoveException();
        }

        private static bool TrySetDispositionGroupsForStandardModuleSource(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, out Dictionary<MoveDisposition, List<Declaration>> dispositions )
        {
            dispositions = EmptyDispositions();

            //Any Private declaration (member, field, or constant) that is used by both the Move Participants
            //and NonParticipants MUST be retained in the Source module.  If a Selected Element (which HAS to move)
            //directly references a declaration that MUST be retained - the move cannot be executed
            if (SelectedElementsDirectlyReferenceNonExclusivePrivateElement(moveGroups))
            {
                return false;
            }

            var support = new Dictionary<RequiredGroup, List<Declaration>>()
            {
                [RequiredGroup.PrivateMove] = new List<Declaration>(),
                [RequiredGroup.PrivateRetain] = PrivateDeclarationsThatMustBeRetained(moveGroups),
                [RequiredGroup.PublicMove] = new List<Declaration>(),
                [RequiredGroup.PublicRetain] = moveGroups.Declarations(MoveGroup.Support_Public).ToList(),
            };

            foreach (var selectedMoveableMemberSetDependency in moveGroups.ToMoveableMemberSets(DependenciesRequiredToMove(moveGroups)))
            {
                //All directly referenced Private declarations of a Selected element must be moved.
                //Otherwise, the Selected element cannot be moved
                if (selectedMoveableMemberSetDependency.HasPrivateAccessibility)
                {
                    support[RequiredGroup.PrivateMove].AddRange(selectedMoveableMemberSetDependency.Members);
                }
            }

            var publicSupportMoveables = moveGroups.MoveableMemberSets(MoveGroup.Support_Public);
            for (var idx = 0; idx < publicSupportMoveables.Count; idx++)
            {
                var moveable = publicSupportMoveables.ElementAt(idx);
                if (!support[RequiredGroup.PublicRetain].Contains(moveable.Member))
                {
                    continue;
                }

                //If a Public support members references a Private support declaration that 'must move'
                //then the Public support member 'must move' as well.
                if (IsMustMovePublicSupport(moveable, support[RequiredGroup.PrivateMove], out var newPrivateMustMoveSupport))
                {
                    support[RequiredGroup.PublicMove].AddRange(moveable.Members);

                    support[RequiredGroup.PublicRetain] = support[RequiredGroup.PublicRetain]
                                                                            .Except(moveable.Members)
                                                                            .ToList();

                    support[RequiredGroup.PrivateMove] = support[RequiredGroup.PrivateMove]
                                                                            .Concat(newPrivateMustMoveSupport)
                                                                            .Distinct()
                                                                            .ToList();

                    //Need to work from the start of the list again to see if the added private support
                    //dependencies forces a move of any other MoveableMembers in the 'Retain' collection 
                    idx = -1;
                }
            }

            var retainedPrivateSupportCandidates = moveGroups.Declarations(MoveGroup.Support_Private).Except(support[RequiredGroup.PrivateMove]);

            var privateDependenciesOfRetainedPublicSupportMembers = retainedPrivateSupportCandidates.AllReferences()
                            .Where(rf => support[RequiredGroup.PublicRetain].Contains(rf.ParentScoping))
                            .Select(rf => rf.Declaration);

            support[RequiredGroup.PrivateRetain].AddRange(privateDependenciesOfRetainedPublicSupportMembers);

            var privateExclusiveSupportMoveableMemberSets = moveGroups.MoveableMemberSets(MoveGroup.Support_Exclusive)
                                .Where(p => p.HasPrivateAccessibility)
                                .Except(moveGroups.ToMoveableMemberSets(support[RequiredGroup.PrivateRetain]));

            //if Private exclusive support members (which must move) have
            //a declaration in common with the Private declarations that must be retained in the Source
            //Module...the move is not executable.
            if (privateExclusiveSupportMoveableMemberSets.SelectMany(mm => mm.DirectDependencies)
                .Intersect(support[RequiredGroup.PrivateRetain]).Any())
            {
                return false;
            }

            support[RequiredGroup.PrivateMove].AddRange(privateExclusiveSupportMoveableMemberSets.SelectMany(mm => mm.Members));

            foreach (var key in support.Keys.ToList())
            {
                support[key] = support[key].Distinct().ToList();
            }

            //Final check to see that all 'binning' has not resulted in overlaps.  If there are overlaps,
            //the strategy cannot fully resolve the scenario and execute the move 
            if (support[RequiredGroup.PrivateMove].Intersect(support[RequiredGroup.PrivateRetain]).Any()
                || support[RequiredGroup.PublicMove].Intersect(support[RequiredGroup.PublicRetain]).Any())
            {
                return false;
            }

            if (CreatesEnumOrEnumMemberNameConflict(model, support[RequiredGroup.PrivateMove]))
            {
                return false;
            }

            dispositions[MoveDisposition.Move] = (moveGroups.Declarations(MoveGroup.Selected)
                                                                        .Concat(support[RequiredGroup.PublicMove])
                                                                        .Concat(support[RequiredGroup.PrivateMove])).ToList();

            dispositions[MoveDisposition.Retain] = (moveGroups.Declarations(MoveGroup.AllParticipants)
                                                                        .Except(dispositions[MoveDisposition.Move])).ToList();
            return true;
        }

        private static bool TrySetDispositionGroupsForClassModuleSource(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, out Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            dispositions = EmptyDispositions();

            var participatingMembers = moveGroups.MoveableMemberSets(MoveGroup.AllParticipants)
                        .Where(d => d.Member.IsMember())
                        .SelectMany(pm => pm.Members);

            if (ParticipantsRaiseAnEvent(participatingMembers)
                || ParticipantsIncludeAnEventHandler(participatingMembers, model)
                || ParticipantsIncludeAnInterfaceImplementingMember(participatingMembers, model))
            {
                return false;
            }

            if (moveGroups.Declarations(MoveGroup.Support_NonExclusive).Any()) { return false; }

            //Strategy does not move Public Fields of ObjectTypes.
            //As a Public Field of a StandardModule, we lose control of instantiation logic
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && nm.IsObject)) { return false; }


            //If there are external references to the selected element(s) and the source is a 
            //ClassModule or  UserForm, then they are not be supported by this strategy.
            if (model.SelectedDeclarations.AllReferences().Any(rf => rf.QualifiedModuleName != model.Source.QualifiedModuleName)) { return false; }

            dispositions[MoveDisposition.Move] = (moveGroups.Declarations(MoveGroup.Selected)
                .Concat(moveGroups.Declarations(MoveGroup.Support_Exclusive))).ToList();

            dispositions[MoveDisposition.Retain] = (moveGroups.Declarations(MoveGroup.AllParticipants)
                                                        .Except(dispositions[MoveDisposition.Move])).ToList();
            return true;
        }

         //ClearMoveNameConflicts renames members that, when moved to the Destination module as-is, will
        //result in a name conflict.  The renaming occurs in both the moveMemberRewriteSession and
        //in the scratchPadSession so that moved blocks of code generated from the scrathpad already have
        //the new names loaded to the rewriter.
        private static void ClearMoveNameConflicts(MoveMemberModel model, RenameCodeDefinedIdentifierRefactoringAction renameAction, IRewriteSession moveMemberRewriteSession, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            foreach (var moveable in model.MoveableMembers)
            {
                moveable.MovedIdentifierName = moveable.IdentifierName;
            }

            SetNonConflictIdentifierNames(model, dispositions[MoveDisposition.Move]);

            foreach (var moveable in model.MoveableMembers)
            {
                RenameMoveableMemberSet(moveable, renameAction, moveMemberRewriteSession);
                RenameMoveableMemberSet(moveable, renameAction, scratchPadSession);
            }
        }

        private static void RenameMoveableMemberSet(IMoveableMemberSet moveableMemberSet, RenameCodeDefinedIdentifierRefactoringAction renameAction, IRewriteSession rewriteSession)
        {
            foreach (var member in moveableMemberSet.Members)
            {
                if (member.IdentifierName.IsEquivalentVBAIdentifierTo(moveableMemberSet.MovedIdentifierName))
                {
                    continue;
                }

                var renameModel = new RenameModel(member)
                {
                    NewName = moveableMemberSet.MovedIdentifierName
                };

                renameAction.Refactor(renameModel, rewriteSession);
            }
        }

        private static IMovedContentProvider LoadMovedContentProvider(IMovedContentProvider contentProvider, MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            foreach (var element in dispositions[MoveDisposition.Move])
            {
                if (element.IsMember())
                {
                    var memberCodeBlock = CreateMovedMemberCodeBlock(model, moveGroups, element, scratchPadSession, dispositions);
                    contentProvider.AddMethodDeclaration(memberCodeBlock);
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

        private static void SetNonConflictIdentifierNames(MoveMemberModel model, IEnumerable<Declaration> membersToMove)
        {
            var conflictingMoveables = new List<IMoveableMemberSet>();
            var moveableMembers = membersToMove.Select(mtm => model.MoveableMemberSetByName(mtm.IdentifierName)).Distinct();

            foreach (var moveableMember in moveableMembers)
            {
                var movingEnumMembers = moveableMember.IsEnumeration
                    ? model.Source.ModuleDeclarations.Where(d => d.DeclarationType.Equals(DeclarationType.EnumerationMember) && moveableMember.Member.Equals(d.ParentDeclaration))
                    : null;

                if (DeclarationMoveCreatesNameConflict(moveableMember.Member, model.Destination.ModuleDeclarations, movingEnumMembers))
                {
                    conflictingMoveables.Add(moveableMember);
                }
            }

            var allPotentialConflictNames = model.Destination.ModuleDeclarations
                .Select(m => m.IdentifierName).ToList();

            foreach (var conflictMoveable in conflictingMoveables)
            {
                if (!TryCreateNonConflictIdentifier(conflictMoveable.IdentifierName, allPotentialConflictNames, out var identifier))
                {
                    throw new MoveMemberUnsupportedMoveException(conflictMoveable?.Member);
                }

                conflictMoveable.MovedIdentifierName = identifier;
                allPotentialConflictNames.Add(identifier);
            }
        }

        private static bool TryCreateNonConflictIdentifier(string originalIdentifier, IEnumerable<string> allPotentialConflictNames, out string identifier)
        {
            identifier = originalIdentifier;

            var guard = 0;
            var maxIterations = 50;
            while (guard++ < maxIterations && allPotentialConflictNames.Contains(identifier))
            {
                identifier = identifier.IncrementIdentifier();
            }
            return guard < maxIterations;
        }

        private static void ModifyRetainedReferencesToMovedMembers(IRewriteSession rewriteSession, MoveMemberModel model, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var renamableReferences = RenameableReferences(dispositions[MoveDisposition.Move], model.Source.QualifiedModuleName);
            var retainedReferencesToModuleQualify = renamableReferences.Where(rf => !dispositions[MoveDisposition.Move].Contains(rf.ParentScoping));

            var moveableConstants = model.MoveableMembers.Where(mm => mm.Member.IsConstant());
            var directReferencesOfMovedConstants = new List<IdentifierReference>();
            foreach (var constant in moveableConstants)
            {
                if (dispositions[MoveDisposition.Move].Contains(constant.Member))
                {
                    directReferencesOfMovedConstants.AddRange(constant.DirectReferences);
                    retainedReferencesToModuleQualify = retainedReferencesToModuleQualify.Except(constant.DirectReferences);
                }
            }

            var moveableFields = model.MoveableMembers.Where(mm => mm.Member.IsMemberVariable());
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

        private static string InsertDestinationContent(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, IRewriteSession moveMemberRewriteSession, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions, IMovedContentProvider movedContentProvider)
        {
            movedContentProvider = LoadMovedContentProvider(movedContentProvider, model, moveGroups, scratchPadSession, dispositions);

            var newContent = movedContentProvider.AsSingleBlock;

            InsertMovedContent(moveMemberRewriteSession, model.Destination, newContent);
            return newContent;
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

        private static void InsertMovedContent(IRewriteSession refactoringRewriteSession, IMoveDestinationEndpoint destination, string movedContent)
        {
            if (!destination.IsExistingModule(out var module))
            {
                throw new MoveMemberUnsupportedMoveException();
            }

            var destinationRewriter = refactoringRewriteSession.CheckOutModuleRewriter(module.QualifiedModuleName);

            if (destination.TryGetCodeSectionStartIndex(out var codeSectionStartIndex))
            {
                destinationRewriter.InsertBefore(codeSectionStartIndex, $"{movedContent}{Environment.NewLine}{Environment.NewLine}");
            }
            else
            {
                destinationRewriter.InsertAtEndOfFile($"{Environment.NewLine}{Environment.NewLine}{movedContent}");
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
                    ? member.Accessibility.TokenString()
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

        private static void UpdateReferencesToMovedMembersInNonEndpointModules(MoveMemberModel model, IRewriteSession rewriteSession, Dictionary<MoveDisposition, List<Declaration>>  dispositions)
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

        private static string AddDestinationModuleQualification(MoveMemberModel model, IdentifierReference identifierReference, IEnumerable<Declaration> retain)
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

        private static bool IsOnlyReferencedByMovedMethods(Declaration element, IEnumerable<Declaration> move)
            => element.References.All(rf => move.Where(m => m.IsMember()).Contains(rf.ParentScoping));

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

        private static bool ParticipantsIncludeAnInterfaceImplementingMember(IEnumerable<Declaration> participatingMembers, MoveMemberModel model)
        {
            var interfaceImplementingMembers = model.DeclarationFinderProvider.DeclarationFinder.FindAllInterfaceImplementingMembers()
                .Where(ifm => ifm.QualifiedModuleName == model.Source.QualifiedModuleName);

            return participatingMembers.Intersect(interfaceImplementingMembers).Any();
        }

        private static bool ParticipantsRaiseAnEvent(IEnumerable<Declaration> participatingMembers)
                    => participatingMembers.Where(m => m.Context.GetDescendent<VBAParser.RaiseEventStmtContext>() != null).Any();

        private static bool ParticipantsIncludeAnEventHandler(IEnumerable<Declaration> participatingMembers, MoveMemberModel model)
        {
            var eventHandlers = model.DeclarationFinderProvider.DeclarationFinder.FindEventHandlers()
                .Where(evh => evh.QualifiedModuleName == model.Source.QualifiedModuleName);

            return participatingMembers.Intersect(eventHandlers).Any();
        }

        //If one of the Private declarations that MUST be retained is also a direct
        //dependency of a Selected member (which MUST move), then game-over.
        private static bool SelectedElementsDirectlyReferenceNonExclusivePrivateElement(IMoveMemberGroupsProvider moveGroups)
        {
            return DependenciesRequiredToMove(moveGroups).Intersect(PrivateDeclarationsThatMustBeRetained(moveGroups)).Any();
        }

        //Any Private declaration (member, field, or constant) that is used by both the Participants
        //and the NonParticipants MUST be retained in the Source module
        private static List<Declaration> PrivateDeclarationsThatMustBeRetained(IMoveMemberGroupsProvider moveGroups)
        {
            var allParticipantDependencies = moveGroups.Dependencies(MoveGroup.AllParticipants);
            var nonParticipantDependencies = moveGroups.Dependencies(MoveGroup.NonParticipants);

            return allParticipantDependencies.Intersect(nonParticipantDependencies)
                .Where(d => d.HasPrivateAccessibility()).ToList();
        }

        private static List<Declaration> DependenciesRequiredToMove(IMoveMemberGroupsProvider moveGroups)
            => moveGroups.DirectDependencies(MoveGroup.Selected).Intersect(moveGroups.Declarations(MoveGroup.Support_Private)).ToList();

        private static IEnumerable<IdentifierReference> RenameableReferences(IEnumerable<Declaration> declarations, QualifiedModuleName qmn)
            => RenameableReferencesByQualifiedModuleName(declarations.AllReferences())
                                                            .Where(g => qmn == g.Key)
                                                            .SelectMany(g => g);

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

        private static bool CreatesEnumOrEnumMemberNameConflict(MoveMemberModel model, IEnumerable<Declaration> moveDeclarations)
        {
            //Since these are existing declarations we are moving around, we only need to check
            //moved Private Enumerations and their Members for conflicts
            var privateMoveDeclarations = moveDeclarations.Where(d => d.HasPrivateAccessibility());

            var movingEnums = privateMoveDeclarations.Where(d => d.DeclarationType.Equals(DeclarationType.Enumeration));

            if (!movingEnums.Any()) { return false; }

            foreach (var movingEnum in movingEnums)
            {
                var movingEnumMembers = model.Source.ModuleDeclarations.Where(d => d.DeclarationType.Equals(DeclarationType.EnumerationMember) && movingEnum.Equals(d.ParentDeclaration));
                if (DeclarationMoveCreatesNameConflict(movingEnum, model.Destination.ModuleDeclarations, movingEnumMembers))
                {
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

        private static Dictionary<MoveDisposition, List<Declaration>> EmptyDispositions()
        {
            return new Dictionary<MoveDisposition, List<Declaration>>()
            {
                [MoveDisposition.Move] = new List<Declaration>(),
                [MoveDisposition.Retain] = new List<Declaration>()
            };
        }
    }
}
