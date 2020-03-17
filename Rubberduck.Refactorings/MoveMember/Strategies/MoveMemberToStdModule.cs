using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
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
            PrivateSupportRetain,
            PrivateSupportMove,
            PublicSupportRetain,
            PublicSupportMove,
        }

        private enum MoveDisposition
        {
            Move,
            Retain
        }

        public override bool IsApplicable(MoveMemberModel model)
        {
            //A strategy exists to handle 'nothing selected'
            if (!model.SelectedDeclarations.Any()) { return false; }

            //Note: A StandardModule is the default so... 
            //model.Destination.IsStandardModule returns true when nothing is specified
            if (!model.Destination.IsStandardModule) { return false; }

            //Strategy does not support Private Fields as a SelectedDeclaration
            if (model.SelectedDeclarations.Any(sd => sd.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && sd.HasPrivateAccessibility())) { return false; }

            var moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.MoveableMembers);

            if (model.Source.IsStandardModule)
            {
                return TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out _);
            }

            return TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out _);
        }

        public override bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage)
        {
            nonExecutableMessage = string.Empty;
            
            if (string.IsNullOrEmpty(model?.Destination.ModuleName))
            {
                nonExecutableMessage = MoveMemberResources.UndefinedDestinationModule;
                return false;
            }
            return true;
        }

        public override void RefactorRewrite(MoveMemberModel model, IRewriteSession moveMemberRewriteSession, IRewritingManager rewritingManager, bool asPreview = false)
        {
            if (!asPreview && string.IsNullOrEmpty(model.Destination.ModuleName))
            {
                throw new MoveMemberUnsupportedMoveException();
            }

            var dispositions = new Dictionary<MoveDisposition, List<Declaration>>()
            {
                [MoveDisposition.Move] = new List<Declaration>(),
                [MoveDisposition.Retain] = new List<Declaration>()
            };

            var moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.MoveableMembers);
            if (model.Source.IsStandardModule)
            {
                if (!TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out dispositions))
                {
                    throw new MoveMemberUnsupportedMoveException();
                }
            }
            else
            {
                if (!TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out dispositions))
                {
                    throw new MoveMemberUnsupportedMoveException();
                }
            }

            RemoveDeclarations(moveMemberRewriteSession, dispositions[MoveDisposition.Move]);

            if (model.Destination.IsExistingModule(out var destinationModule))
            {
                var movedContentProvider = model.MoveMemberFactory.CreateMovedContentProvider();
                movedContentProvider = LoadMovedContentProvider(movedContentProvider, model, moveGroups, rewritingManager.CheckOutCodePaneSession(), dispositions);

                var newContent = asPreview
                    ? movedContentProvider.AsSingleBlockWithinDemarcationComments()
                    : movedContentProvider.AsSingleBlock;

                InsertMovedContent(moveMemberRewriteSession, model.Destination, newContent);

                ModifyDestinationExistingReferencesToMovedMembers(destinationModule, moveGroups, moveMemberRewriteSession);
            }

            ModifyRetainedReferencesToMovedMembers(moveMemberRewriteSession, model, moveGroups, dispositions);

            UpdateOtherModuleReferencesToMovedMembers(model, moveGroups, moveMemberRewriteSession);
        }

        public override IMovedContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, IMovedContentProvider contentToMove)
        {
            var dispositions = new Dictionary<MoveDisposition, List<Declaration>>()
            {
                [MoveDisposition.Move] = new List<Declaration>(),
                [MoveDisposition.Retain] = new List<Declaration>()
            };

            var moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.MoveableMembers);
            if (model.Source.IsStandardModule)
            {
                if (!TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out dispositions))
                {
                    throw new MoveMemberUnsupportedMoveException();
                }
            }
            else
            {
                if (!TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out dispositions))
                {
                    throw new MoveMemberUnsupportedMoveException();
                }
            }

            var movedContentProvider = model.MoveMemberFactory.CreateMovedContentProvider();
            return LoadMovedContentProvider(movedContentProvider, model, moveGroups, rewritingManager.CheckOutCodePaneSession(), dispositions);
        }

        private static bool TrySetDispositionGroupsForClassModuleSource(MoveMemberModel model, IMoveGroupsProvider moveGroups, out Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            dispositions = new Dictionary<MoveDisposition, List<Declaration>>()
            {
                [MoveDisposition.Move] = new List<Declaration>(),
                [MoveDisposition.Retain] = new List<Declaration>()
            };

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
            //As a Public Field of a StandardModule, we cannot guarantee that it is a valid instance
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && nm.IsObject)) { return false; }


            //If there are external references to the selected element(s) and the source is a 
            //ClassModule or  UserForm, then they are not be supported by this strategy.
            if (model.SelectedDeclarations.AllReferences().Any(rf => rf.QualifiedModuleName != model.Source.QualifiedModuleName)) { return false; }

            dispositions[MoveDisposition.Move] = moveGroups.Declarations(MoveGroup.Selected)
                .Concat(moveGroups.Declarations(MoveGroup.Support_Exclusive))
                .ToList();

            dispositions[MoveDisposition.Retain] = moveGroups.Declarations(MoveGroup.AllParticipants)
                                                        .Except(dispositions[MoveDisposition.Move]).ToList();
            return true;
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

        private static bool TrySetDispositionGroupsForStandardModuleSource(MoveMemberModel model, IMoveGroupsProvider moveGroups, out Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var emptyDispositions = new Dictionary<MoveDisposition, List<Declaration>>()
            {
                [MoveDisposition.Move] = new List<Declaration>(),
                [MoveDisposition.Retain] = new List<Declaration>()
            };

            dispositions = emptyDispositions;

            //Strategy does not move Public Fields of ObjectTypes.
            //As a Public Field of a StandardModule, we cannot guarantee that it is a valid instance
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && nm.IsObject)) { return false; }

            var requiredGroup = new Dictionary<RequiredGroup, List<Declaration>>()
            {
                [RequiredGroup.PrivateSupportMove] = new List<Declaration>(),
                [RequiredGroup.PrivateSupportRetain] = new List<Declaration>(),
                [RequiredGroup.PublicSupportMove] = new List<Declaration>(),
                [RequiredGroup.PublicSupportRetain] = new List<Declaration>(),
            };

            //This strategy attempts to move only members explicitly selected by the user/caller and
            //retain as many Public support members as possible. 
            requiredGroup[RequiredGroup.PublicSupportRetain] = moveGroups.Declarations(MoveGroup.Support_Public).ToList();

            var allParticipantDependencies = moveGroups.Dependencies(MoveGroup.AllParticipants);
            var nonParticipantDependencies = moveGroups.Dependencies(MoveGroup.NonParticipants);

            //Any Private declaration (member, field, or constant) that is used by both the Participants
            //and the NonParticipants MUST be retained in the Source module
            requiredGroup[RequiredGroup.PrivateSupportRetain] = allParticipantDependencies.Intersect(nonParticipantDependencies)
                .Where(d => d.HasPrivateAccessibility()).ToList();

            var selectedMemberDirectDependencies = moveGroups.MoveableMemberSets(MoveGroup.Selected)
                .SelectMany(mm => mm.DirectDependencies);

            var selectedMemberDependenciesRequiredToMove = selectedMemberDirectDependencies.Intersect(moveGroups.Declarations(MoveGroup.Support_Private));
            if (selectedMemberDependenciesRequiredToMove.Intersect(requiredGroup[RequiredGroup.PrivateSupportRetain]).Any())
            {
                //If one of the Private declarationa that MUST be retained is also direct
                //dependency of a Selected member, then game-over.
                //The move cannot be supportd by this strategy.
                return false;
            }

            var selectedMemberPrivateTypeDependencies = moveGroups.MoveableMemberSets(MoveGroup.Selected)
                .SelectMany(mm => mm.Members.Where(m => m.AsTypeDeclaration?.HasPrivateAccessibility() ?? false));

            if (selectedMemberPrivateTypeDependencies.Any())
            {
                return false;
            }

            var privateExclusiveSupportMoveableMemberSets = moveGroups.MoveableMemberSets(MoveGroup.Support_Exclusive)
                .Where(p => p.HasPrivateAccessibility).ToList();

            var mustMoveDependencySets = moveGroups.ToMoveableMemberSets(selectedMemberDependenciesRequiredToMove);
            foreach (var selectedMoveableMemberSetDependency in mustMoveDependencySets)
            {
                //Add all the Selecteds' Private dependencies to the Required set of declarations that must be moved
                if (selectedMoveableMemberSetDependency.HasPrivateAccessibility)
                {
                    requiredGroup[RequiredGroup.PrivateSupportMove].AddRange(selectedMoveableMemberSetDependency.Members);
                    privateExclusiveSupportMoveableMemberSets.RemoveAll(pe => selectedMoveableMemberSetDependency.IdentifierName.IsEquivalentVBAIdentifierTo(pe.IdentifierName));
                }
            }

            var publicSupportMoveables = moveGroups.MoveableMemberSets(MoveGroup.Support_Public);
            for (var idx = 0; idx < publicSupportMoveables.Count; idx++)
            {
                var moveable = publicSupportMoveables.ElementAt(idx);
                if (!requiredGroup[RequiredGroup.PublicSupportRetain].Contains(moveable.Member))
                {
                    continue;
                }

                //If a Public support members references a Private support declaration that 'must move'
                //then the Public support member 'must move' as well.
                if (IsMustMovePublicSupport(moveable, requiredGroup[RequiredGroup.PrivateSupportMove], out var newPrivateMustMoveSupport))
                {
                    requiredGroup[RequiredGroup.PublicSupportMove].AddRange(moveable.Members);
                    requiredGroup[RequiredGroup.PublicSupportRetain] = requiredGroup[RequiredGroup.PublicSupportRetain].Except(moveable.Members).ToList();

                    requiredGroup[RequiredGroup.PrivateSupportMove] = requiredGroup[RequiredGroup.PrivateSupportMove].Concat(newPrivateMustMoveSupport).Distinct().ToList();
                    //Need to work from the start of the list again to see if the added private support
                    //dependencies forces a move of any other MoveableMembers in the 'Retain' collection 
                    idx = -1;
                }
            }

            var retainedPrivateSupportCandidates = moveGroups.Declarations(MoveGroup.Support_Private).Except(requiredGroup[RequiredGroup.PrivateSupportMove]);

            var privateDependenciesOfRetainedPublicSupportMembers = retainedPrivateSupportCandidates.AllReferences()
                            .Where(rf => requiredGroup[RequiredGroup.PublicSupportRetain].Contains(rf.ParentScoping))
                            .Select(rf => rf.Declaration);


            requiredGroup[RequiredGroup.PrivateSupportRetain].AddRange(privateDependenciesOfRetainedPublicSupportMembers);

            privateExclusiveSupportMoveableMemberSets = privateExclusiveSupportMoveableMemberSets
                            .Except(moveGroups.ToMoveableMemberSets(requiredGroup[RequiredGroup.PrivateSupportRetain])).ToList();

            //if Private exclusive support members (which must move with the selected declarations) have
            //a declaration in common with the Private declarations that must be retained in the Source
            //Module...no can do.
            if (privateExclusiveSupportMoveableMemberSets.SelectMany(mm => mm.DirectDependencies)
                .Intersect(requiredGroup[RequiredGroup.PrivateSupportRetain]).Any())
            {
                return false;
            }

            //And if the Private support declarations that are required to move have any common
            //declarations with the Private support declarations required to be retained, then this is 
            //also a non-executable move
            if (requiredGroup[RequiredGroup.PrivateSupportMove].Intersect(requiredGroup[RequiredGroup.PrivateSupportRetain]).Any())
            {
                return false;
            }

            dispositions[MoveDisposition.Move] = moveGroups.Declarations(MoveGroup.Selected)
                                    .Concat(privateExclusiveSupportMoveableMemberSets.SelectMany(mm => mm.Members))
                                    .Concat(requiredGroup[RequiredGroup.PublicSupportMove])
                                    .Concat(requiredGroup[RequiredGroup.PrivateSupportMove])
                                    .Except(requiredGroup[RequiredGroup.PublicSupportRetain])
                                    .Except(requiredGroup[RequiredGroup.PrivateSupportRetain])
                                    .Distinct()
                                    .ToList();

            dispositions[MoveDisposition.Retain] = moveGroups.Declarations(MoveGroup.AllParticipants)
                                                        .Except(dispositions[MoveDisposition.Move]).ToList();
            return true;
        }

        private static IMovedContentProvider LoadMovedContentProvider(IMovedContentProvider contentProvider, MoveMemberModel model, IMoveGroupsProvider moveGroups, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            ModifyMovedIdentifiersToAvoidDestinationNameConflicts(model, dispositions[MoveDisposition.Move]);

            var rewriter = scratchPadSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);

            foreach (var element in dispositions[MoveDisposition.Move])
            {
                if (element.IsMember())
                {
                    var memberCodeBlock = CreateMovedMemberCodeBlock(model, moveGroups, element, rewriter, dispositions);
                    contentProvider.AddMethodDeclaration(memberCodeBlock);
                    continue;
                }

                var nonMembercodeBlock = CreateMovedNonMemberCodeBlock(model, moveGroups, element, rewriter, dispositions);
                contentProvider.AddFieldOrConstantDeclaration(nonMembercodeBlock);
            }

            return contentProvider;
        }

        private static void ModifyMovedIdentifiersToAvoidDestinationNameConflicts(MoveMemberModel model, IEnumerable<Declaration> copyAndDelete)
        {
            if (!model.Destination.IsExistingModule(out var destination)) { return; }

            var allPotentialConflictNames = model.Destination.ModuleDeclarations
                .Where(d => !NeverCauseNameConflictTypes.Contains(d.DeclarationType))
                .Select(m => m.IdentifierName).ToList();

            var conflictingIdentifiers = copyAndDelete
                        .Where(m => allPotentialConflictNames.Contains(m.IdentifierName))
                        .Select(m => m.IdentifierName)
                        .Distinct();

            foreach (var conflictIdentifier in conflictingIdentifiers)
            {
                var moveableMemberSet = model.MoveableMemberSetByName(conflictIdentifier);

                if (!TryCreateNonConflictIdentifier(conflictIdentifier, allPotentialConflictNames, out var identifier))
                {
                    throw new MoveMemberUnsupportedMoveException(moveableMemberSet?.Member);
                }

                moveableMemberSet.MovedIdentifierName = identifier;
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

        private static void ModifyRetainedReferencesToMovedMembers(IRewriteSession rewriteSession, MoveMemberModel model, IMoveGroupsProvider moveGroups, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var retainedReferencesToModuleQualify = RenameableReferences(moveGroups.Declarations(MoveGroup.Selected), model.Source.QualifiedModuleName)
                .Where(rf => !dispositions[MoveDisposition.Move].Contains(rf.ParentScoping));

            if (retainedReferencesToModuleQualify.Any())
            {
                var sourceRewriter = rewriteSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);
                foreach (var rf in retainedReferencesToModuleQualify)
                {
                    sourceRewriter.Replace(rf.Context, AddDestinationModuleQualification(model, rf, dispositions[MoveDisposition.Retain]));
                }
            }
        }

        private static void ModifyDestinationExistingReferencesToMovedMembers(Declaration destination, IMoveGroupsProvider moveGroups, IRewriteSession rewriteSession)
        {
            var destinationReferencesToMovedMembers = moveGroups.Declarations(MoveGroup.Selected).AllReferences()
                .Where(rf => rf.QualifiedModuleName == destination.QualifiedModuleName);

            if (destinationReferencesToMovedMembers.Any())
            {
                var destinationRewriter = rewriteSession.CheckOutModuleRewriter(destination.QualifiedModuleName);

                destinationRewriter.RemoveMemberAccess(destinationReferencesToMovedMembers);

                destinationRewriter.RemoveWithMemberAccess(destinationReferencesToMovedMembers);
            }
        }

        private static IModuleRewriter InsertMovedContent(IRewriteSession refactoringRewriteSession, IMoveDestinationModuleProxy destination, string movedContent)
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

            return destinationRewriter;
        }

        private static string CreateMovedMemberCodeBlock(MoveMemberModel model, IMoveGroupsProvider moveGroups, Declaration member, IModuleRewriter rewriter, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            Debug.Assert(member.IsMember());

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

            var moveableMemberSet = model.MoveableMemberSetByName(member.IdentifierName);
            SetMovedMemberIdentifier(moveableMemberSet, member, rewriter);

            var otherMoveParticipantReferencesRelatedToMember = moveGroups.Declarations(MoveGroup.Support_Exclusive)
                                    .Where(esd => !esd.IsMember()).AllReferences()
                                    .Where(rf => rf.ParentScoping == member);

            foreach (var rf in otherMoveParticipantReferencesRelatedToMember)
            {
                rewriter.Rename(rf, model.MoveableMemberSetByName(rf.IdentifierName).MovedIdentifierName);
            }

            if (model.Source.IsStandardModule)
            {
                rewriter = AddSourceModuleQualificationToMovedReferences(member, model.Source.ModuleName, rewriter, dispositions);
            }

            var destinationMemberAccessReferencesToMovedMembers = model.Destination.ModuleDeclarations
                .AllReferences().Where(rf => rf.ParentScoping == member);

            rewriter.RemoveMemberAccess(destinationMemberAccessReferencesToMovedMembers);
            return rewriter.GetText(member);
        }

        private static string CreateMovedNonMemberCodeBlock(MoveMemberModel model, IMoveGroupsProvider moveGroups, Declaration nonMember, IModuleRewriter rewriter, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            Debug.Assert(!nonMember.IsMember());

            var visibility = IsOnlyReferencedByMovedMethods(nonMember, dispositions[MoveDisposition.Move]) ? nonMember.Accessibility.TokenString() : Tokens.Public;

            var moveableMemberSet = model.MoveableMemberSetByName(nonMember.IdentifierName);
            SetMovedMemberIdentifier(moveableMemberSet, nonMember, rewriter);

            if (model.Source.IsStandardModule && nonMember.IsConstant())
            {
                foreach (var rf in moveableMemberSet.DirectReferences)
                {
                    rewriter.Replace(rf.Context, $"{model.Source.ModuleName}.{rf.IdentifierName}");
                }
            }

            return nonMember.IsVariable()
                ? $"{visibility} {rewriter.GetText(nonMember)}"
                : $"{visibility} {Tokens.Const} {rewriter.GetText(nonMember)}";
        }

        private static void SetMovedMemberIdentifier(IMoveableMemberSet membersRelatedByName, Declaration declaration, IModuleRewriter rewriter)
        {
            if (!membersRelatedByName.RetainsOriginalIdentifier)
            {
                rewriter.Rename(declaration, membersRelatedByName.MovedIdentifierName);
            }

            if (!declaration.IsMember()) { return; }

            var renameableReferences = RenameableReferences(declaration, declaration.QualifiedModuleName);

            foreach (var rf in renameableReferences)
            {
                if (!membersRelatedByName.RetainsOriginalIdentifier)
                {
                    rewriter.Rename(rf, membersRelatedByName.MovedIdentifierName);
                }
            }
        }

        private static IModuleRewriter AddSourceModuleQualificationToMovedReferences(Declaration member, string sourceModuleName, IModuleRewriter tempRewriter, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var retainedPublicDeclarations = dispositions[MoveDisposition.Retain].Where(m => !m.HasPrivateAccessibility());
            if (retainedPublicDeclarations.Any())
            {
                var destinationRefs = retainedPublicDeclarations.AllReferences().Where(rf => dispositions[MoveDisposition.Move].Contains(rf.ParentScoping));
                foreach (var rf in destinationRefs)
                {
                    tempRewriter.Replace(rf.Context, $"{sourceModuleName}.{rf.IdentifierName}");
                }
            }
            return tempRewriter;
        }

        private static void UpdateOtherModuleReferencesToMovedMembers(MoveMemberModel model, IMoveGroupsProvider moveGroups, IRewriteSession rewriteSession)
        {
            var endpointQMNs = new List<QualifiedModuleName>() { model.Source.QualifiedModuleName };
            if (model.Destination.IsExistingModule(out var destination))
            {
                endpointQMNs.Add(destination.QualifiedModuleName);
            }

            var qmnToReferenceGroups
                    = RenameableReferencesByQualifiedModuleName(moveGroups.Declarations(MoveGroup.Selected).AllReferences())
                            .Where(qmn => !endpointQMNs.Contains(qmn.Key));

            foreach (var referenceGroup in qmnToReferenceGroups)
            {
                var moduleRewriter = rewriteSession.CheckOutModuleRewriter(referenceGroup.Key);

                var idRefMemberAccessExpressionContextPairs = referenceGroup.Where(rf => rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)
                        .Select(rf => (rf, rf.Context.Parent as VBAParser.MemberAccessExprContext));

                var destinationModuleName = model.Destination.ModuleName;
                foreach (var (IdRef, MemberAccessExpressionContext) in idRefMemberAccessExpressionContextPairs)
                {
                    if (!model.MoveableMemberSetByName(IdRef.IdentifierName).RetainsOriginalIdentifier)
                    {
                        moduleRewriter.Replace(IdRef.Context, model.MoveableMemberSetByName(IdRef.IdentifierName).MovedIdentifierName);
                    }
                    moduleRewriter.Replace(MemberAccessExpressionContext.lExpression(), destinationModuleName);
                }

                var idRefWithMemberAccessExprContextPairs = referenceGroup.Where(rf => rf.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        .Select(rf => (rf, rf.Context.Parent as VBAParser.WithMemberAccessExprContext));

                foreach (var (IdRef, withMemberAccessExprContext) in idRefWithMemberAccessExprContextPairs)
                {
                    if (!model.MoveableMemberSetByName(IdRef.IdentifierName).RetainsOriginalIdentifier)
                    {
                        moduleRewriter.Replace(IdRef.Context, model.MoveableMemberSetByName(IdRef.IdentifierName).MovedIdentifierName);
                    }
                    moduleRewriter.InsertBefore(withMemberAccessExprContext.Start.TokenIndex, destinationModuleName);
                }

                var nonQualifiedReferences = referenceGroup.Where(rf => !(rf.Context.Parent is VBAParser.WithMemberAccessExprContext
                    || (rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)));

                foreach (var rf in nonQualifiedReferences)
                {
                    moduleRewriter.Replace(rf.Context, $"{destinationModuleName}.{model.MoveableMemberSetByName(rf.Declaration.IdentifierName).MovedIdentifierName}");
                }
            }
        }

        private static string AddDestinationModuleQualification(MoveMemberModel model, IdentifierReference identifierReference, IEnumerable<Declaration> retain)
        {
            if (NeverAddMemberAccessTypes.Contains(identifierReference.Declaration.DeclarationType))
            {
                return identifierReference.IdentifierName;
            }

            if (identifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Function)
                && identifierReference.IsAssignment)
            {
                return identifierReference.IdentifierName;
            }

            if (retain.Contains(identifierReference.Declaration))
            {
                return identifierReference.IdentifierName;
            }
            return $"{model.Destination.ModuleName}.{identifierReference.IdentifierName}";
        }

        private static bool IsOnlyReferencedByMovedMethods(Declaration element, IEnumerable<Declaration> copyAndDelete)
            => element.References.All(rf => copyAndDelete.Where(m => m.IsMember()).Contains(rf.ParentScoping));

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


        private static IEnumerable<IdentifierReference> RenameableReferences(Declaration declaration, QualifiedModuleName qmn)
            => RenameableReferences(new Declaration[] { declaration }, qmn);

        private static IEnumerable<IdentifierReference> RenameableReferences(IEnumerable<Declaration> declarations, QualifiedModuleName qmn)
            => RenameableReferencesByQualifiedModuleName(declarations.AllReferences())
                                                            .Where(g => qmn == g.Key)
                                                            .SelectMany(g => g);

        private static IEnumerable<IGrouping<QualifiedModuleName, IdentifierReference>> RenameableReferencesByQualifiedModuleName(IEnumerable<IdentifierReference> references)
        {
            //The filter used by RenameRefactoring
            var modules = references
                .Where(reference =>
                    reference.Context.GetText() != Tokens.Me
                    && !reference.IsArrayAccess
                    && !reference.IsDefaultMemberAccess)
                .GroupBy(r => r.QualifiedModuleName);

            return modules;
        }

        //5.3.1.6 Each<subroutine-declaration> and<function-declaration> must have a procedure 
        //name that is different from any other module variable name, module constant name, 
        //enum member name, or procedure name that is defined within the same module.
        private static List<DeclarationType> NeverCauseNameConflictTypes = new List<DeclarationType>()
        {
            DeclarationType.Project,
            DeclarationType.ProceduralModule,
            DeclarationType.ClassModule,
            DeclarationType.Parameter,
            DeclarationType.Enumeration,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember
        };

        private static List<DeclarationType> NeverAddMemberAccessTypes = new List<DeclarationType>()
        {
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember
        };
    }
}
