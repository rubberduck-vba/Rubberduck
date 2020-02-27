using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToStdModule : IMoveMemberRefactoringStrategy
    {
        public MoveMemberToStdModule()
        {
            _isExecutable = false;
        }

        private MoveMemberModel _model;
        private IRewritingManager _rewritingManager;
        private IMoveGroupsProvider _moveGroups;
        private List<Declaration> _copyAndDelete;
        private List<Declaration> _retain;

        public bool IsApplicable(MoveMemberModel model)
        {
            _isExecutable = !string.IsNullOrEmpty(model.Destination.ModuleName);

            //A strategy exists to handle 'nothing selected'...it's not this one
            if (!model.SelectedDeclarations.Any()) { return false; }

            //Note: A StandardModule is the default so... 
            //model.Destination.IsStandardModule returns true when nothing is specified
            if (!model.Destination.IsStandardModule) { return false; }

            //Strategy does not support Private Fields as a SelectedDeclaration
            if (model.SelectedDeclarations.Any(sd => sd.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && sd.HasPrivateAccessibility())) { return false; }

            var moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.MoveableMembers);


            if (!model.Source.IsStandardModule)
            {
                if (moveGroups.NonExclusiveSupportDeclarations.Any()) { return false; }

                //Strategy does not move Public Fields of ObjectTypes from ClassModules.
                //As a simple Fields, we could no longer guarantee that it is a valid instance
                if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                                        && nm.IsObject)) { return false; }

                
                 //If there are external references to the selected element(s) and the source is a 
                 //ClassModule or  UserForm, then they are not be supported by this strategy            
                var externalMemberRefs = model.SelectedDeclarations.AllReferences().Where(rf => rf.QualifiedModuleName != model.Source.QualifiedModuleName);
                return !externalMemberRefs.Any();
            }

            if (!HasUnMoveableSupportContent(moveGroups, out _)) { return true; }

            return PublicSupportMembersProvideUnMoveableDeclarationsAccess(moveGroups);
        }

        private bool _isExecutable;
        public bool IsAnExecutableScenario(out string nonExecutableMessage)
        {
            nonExecutableMessage = string.Empty;
            if (!_isExecutable)
            {
                if (_model != null)
                {
                    if (string.IsNullOrEmpty(_model.Destination.ModuleName))
                    {
                        nonExecutableMessage = MoveMemberResources.UndefinedDestinationModule;
                    }
                }
            }
            return _isExecutable;
        }

        public IMoveMemberRewriteSession RefactorRewrite(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove, bool asPreview = false)
        {
            if (!asPreview && string.IsNullOrEmpty(model.Destination.ModuleName))
            {
                throw new MoveMemberUnsupportedMoveException();
            }

            InitializeStrategyMembers(model, rewritingManager);

            contentToMove = CreateMovedContentBlock(model, contentToMove, rewritingManager.CheckOutCodePaneSession());
#if DEBUG
            var preModify = GetContentState(model);
#endif
            var moveMemberRewriteSession = model.MoveMemberFactory.CreateMoveMemberRewriteSession(rewritingManager.CheckOutCodePaneSession());

            moveMemberRewriteSession.Remove(_copyAndDelete);

            ModifyRetainedReferencesToMovedMembers(moveMemberRewriteSession);
#if DEBUG
            var postSourceModify = GetContentState(model);
#endif
            if (model.Destination.IsExistingModule(out var destinationModule))
            {
                ModifyDestinationExistingReferencesToMovedMembers(destinationModule, moveMemberRewriteSession);

                var newContent = asPreview
                    ? contentToMove.AsSingleBlockWithinDemarcationComments()
                    : contentToMove.AsSingleBlock;

                InsertMovedContent(moveMemberRewriteSession, destinationModule, newContent);
            }
#if DEBUG
            var postDestinationModify = GetContentState(model);
#endif
            UpdateNonEndpointModuleReferencesToMovedMembers(model, moveMemberRewriteSession);
#if DEBUG
            var finalState = GetContentState(model);
#endif
            return moveMemberRewriteSession;
        }

        public INewContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove)
        {
            InitializeStrategyMembers(model, rewritingManager);
            return CreateMovedContentBlock(model, contentToMove, rewritingManager.CheckOutCodePaneSession());
        }

        private INewContentProvider CreateMovedContentBlock(MoveMemberModel model, INewContentProvider contentToMove, IExecutableRewriteSession scratchPadSession)
        {
            ModifyMovedIdentifiersToAvoidDestinationNameConflicts(model, _copyAndDelete);

            var movedContentRewriter = scratchPadSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);

            foreach (var element in _copyAndDelete)
            {
                var codeBlock = CreateMovedElementCodeBlock(model, element, movedContentRewriter);
                if (element.IsMember())
                {
                    contentToMove.AddMethod(codeBlock);
                    continue;
                }
                contentToMove.AddFieldOrConstantDeclaration(codeBlock);
            }

            return contentToMove;
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

        private void ModifyRetainedReferencesToMovedMembers(IMoveMemberRewriteSession rewriteSession)
        {
            var retainedReferencesToModuleQualify = RenameableReferencesByQualifiedModuleName(_moveGroups.Selected.AllReferences())
                .Where(g => g.Key == _model.Source.QualifiedModuleName)
                .SelectMany(grp => grp)
                .Where(rf => !_copyAndDelete.Contains(rf.ParentScoping));

            if (retainedReferencesToModuleQualify.Any())
            {
                var sourceRewriter = rewriteSession.CheckOutModuleRewriter(_model.Source.QualifiedModuleName);
                foreach (var rf in retainedReferencesToModuleQualify)
                {
                    sourceRewriter.Replace(rf.Context, AddDestinationModuleQualification(rf));
                }
            }
        }

        private void ModifyDestinationExistingReferencesToMovedMembers(Declaration destination, IMoveMemberRewriteSession rewriteSession)
        {
            var destinationReferencesToMovedMembers = _moveGroups.Selected.AllReferences()
                .Where(rf => rf.QualifiedModuleName == destination.QualifiedModuleName);

            if (destinationReferencesToMovedMembers.Any())
            {
                var destinationRewriter = rewriteSession.CheckOutModuleRewriter(destination.QualifiedModuleName);

                destinationRewriter.RemoveMemberAccess(destinationReferencesToMovedMembers);

                destinationRewriter.RemoveWithMemberAccess(destinationReferencesToMovedMembers);
            }
        }

        private IModuleRewriter InsertMovedContent(IMoveMemberRewriteSession refactoringRewriteSession, Declaration destination, string movedContent)
        {
            var destinationRewriter = refactoringRewriteSession.CheckOutModuleRewriter(destination.QualifiedModuleName);
            if (_model.Destination.TryGetCodeSectionStartIndex(out var codeSectionStartIndex))
            {
                destinationRewriter.InsertBefore(codeSectionStartIndex, $"{movedContent}{Environment.NewLine}{Environment.NewLine}");
            }
            else
            {
                destinationRewriter.InsertAtEndOfFile($"{Environment.NewLine}{Environment.NewLine}{movedContent}");
            }

            return destinationRewriter;
        }

        private string CreateMovedElementCodeBlock(MoveMemberModel model, Declaration member, IModuleRewriter rewriter)
        {
            var membersRelatedByName = model.MoveableMemberSetByName(member.IdentifierName);

            if (member.IsVariable() || member.IsConstant())
            {
                var visibility  = IsOnlyReferencedByMovedMethods(member, _copyAndDelete)  ? member.Accessibility.TokenString() : Tokens.Public;

                SetMovedMemberIdentifier(membersRelatedByName, member, rewriter);

                return member.IsVariable()
                    ? $"{visibility} {rewriter.GetText(member)}"
                    : $"{visibility} {Tokens.Const} {rewriter.GetText(member)}";
            }

            if (member is IParameterizedDeclaration paramDeclaration)
            {
                var argListContext = member.Context.GetDescendent<VBAParser.ArgListContext>();
                rewriter.Replace(argListContext, paramDeclaration.BuildFullyDefinedArgumentList());
            }

            if (_moveGroups.Selected.Contains(member))
            {
                var accessibility = IsOnlyReferencedByMovedMethods(member, _copyAndDelete)
                    ? member.Accessibility.TokenString() 
                    : Tokens.Public;

                rewriter.SetMemberAccessibility(member, accessibility);
            }

            SetMovedMemberIdentifier(membersRelatedByName, member, rewriter);
            var otherMoveParticipantReferencesRelatedToMember = _moveGroups.ExclusiveSupportDeclarations.Where(esd => !esd.IsMember()).AllReferences()
                .Where(rf => rf.ParentScoping == member);

            foreach (var rf in otherMoveParticipantReferencesRelatedToMember)
            {
                rewriter.Rename(rf, model.MoveableMemberSetByName(rf.IdentifierName).MovedIdentifierName);
            }

            if (model.Source.IsStandardModule)
            {
                rewriter = AddSourceModuleQualificationToMovedReferences(member, model.Source.ModuleName, rewriter);
            }

            var destinationMemberAccessReferencesToMovedMembers = model.Destination.ModuleDeclarations
                .AllReferences().Where(rf => rf.ParentScoping == member);

            rewriter.RemoveMemberAccess(destinationMemberAccessReferencesToMovedMembers);
            return rewriter.GetText(member);
        }

        private void SetMovedMemberIdentifier(IMoveableMemberSet membersRelatedByName, Declaration declaration, IModuleRewriter rewriter)
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

        private IModuleRewriter AddSourceModuleQualificationToMovedReferences(Declaration member, string sourceModuleName, IModuleRewriter tempRewriter)
        {
            var retainedPublicDeclarations = _retain.Where(m => !m.HasPrivateAccessibility());
            if (retainedPublicDeclarations.Any())
            {
                var destinationRefs = retainedPublicDeclarations.AllReferences().Where(rf => _copyAndDelete.Contains(rf.ParentScoping));
                foreach (var rf in destinationRefs)
                {
                    tempRewriter.Replace(rf.Context, $"{sourceModuleName}.{rf.IdentifierName}");
                }
            }
            return tempRewriter;
        }

        private void UpdateNonEndpointModuleReferencesToMovedMembers(MoveMemberModel model, IMoveMemberRewriteSession rewriteSession)
        {
            var endpointQMNs = new List<QualifiedModuleName>() { model.Source.QualifiedModuleName };
            if (model.Destination.IsExistingModule(out var destination))
            {
                endpointQMNs.Add(destination.QualifiedModuleName);
            }

            var qmnToReferenceGroups 
                    = RenameableReferencesByQualifiedModuleName(_moveGroups.Selected.AllReferences())
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

        private string AddDestinationModuleQualification(IdentifierReference identifierReference)
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

            if (_retain.Contains(identifierReference.Declaration))
            {
                return identifierReference.IdentifierName;
            }
            return $"{_model.Destination.ModuleName}.{identifierReference.IdentifierName}";
        }

        private bool IsOnlyReferencedByMovedMethods(Declaration element, IEnumerable<Declaration> copyAndDelete)
            => element.References.All(rf => copyAndDelete.Where(m => m.IsMember()).Contains(rf.ParentScoping));

#if DEBUG
        private (string Source, string Destination) GetContentState(MoveMemberModel model)
        {
            var debugPreviewSession = _rewritingManager.CheckOutCodePaneSession();
            var sourceRewriter = debugPreviewSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);
            var newSource = sourceRewriter.GetText();
            if (model.Destination.IsExistingModule(out var module))
            {
                var destinationRewriter = debugPreviewSession.CheckOutModuleRewriter(module.QualifiedModuleName);

                var newDestination = destinationRewriter.GetText();
                return (newSource, newDestination);
            }
            return (newSource, string.Empty);
        }
#endif

        private void InitializeStrategyMembers(MoveMemberModel model, IRewritingManager rewritingManager)
        {
            _model = model;

            _rewritingManager = rewritingManager;

            _moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.MoveableMembers);

            var privateSupportToRetainCandidates = new List<Declaration>();
            var publicSupportMembersToRetain = _moveGroups.Declarations(MoveGroups.Support_Public).ToList();

            if (model.Source.IsStandardModule)
            {
                var allParticipantDependencies = _moveGroups.Dependencies(MoveGroups.AllParticipants);
                var nonParticipantDependencies = _moveGroups.Dependencies(MoveGroups.NonParticipants);

                var mustRetainPrivateDeclarations = allParticipantDependencies.Intersect(nonParticipantDependencies)
                    .Where(d => d.HasPrivateAccessibility());

                var selectedMembersDirectDependencies = new List<Declaration>();
                foreach (var selected in _moveGroups[MoveGroups.Selected])
                {
                    selectedMembersDirectDependencies.AddRange(selected.DirectDependencies);
                }

                var privateSupport = _moveGroups.Declarations(MoveGroups.Support).Where(d => d.HasPrivateAccessibility());
                var mustMovePrivateSupport = selectedMembersDirectDependencies.Intersect(privateSupport);

                if (mustMovePrivateSupport.Any(mp => mustRetainPrivateDeclarations.Contains(mp)))
                {
                    //Can't make the move
                    throw new MoveMemberUnsupportedMoveException(_moveGroups.Declarations(MoveGroups.Selected).FirstOrDefault());
                }

                var mustMovePublicSupport = new List<Declaration>();
                var publicSupportMoveables = _moveGroups[MoveGroups.Support_Public];
                for (var idx = 0; idx < publicSupportMoveables.Count; idx++)
                {
                    var moveable = publicSupportMoveables.ElementAt(idx);
                    if (!publicSupportMembersToRetain.Contains(moveable.Member))
                    {
                        continue;
                    }

                    if (IsMustMovePublicSupport(moveable, mustMovePrivateSupport, out var newPrivateMustMoveSupport))
                    {
                        mustMovePrivateSupport = mustMovePrivateSupport.Concat(newPrivateMustMoveSupport).Distinct();
                        publicSupportMembersToRetain = publicSupportMembersToRetain.Except(moveable.Members).ToList();
                        //Need to work from the start of the list again to see if the added private support
                        //dependencies forces a move of any other MoveableMembers in the 'Retain' collection 
                        idx = -1;   
                    }
                }

                var retainCandidates = privateSupport.Except(mustMovePrivateSupport);

                var privateDependenciesOfRetainedPublicSupportMembers = retainCandidates.Intersect(_moveGroups.Dependencies(MoveGroups.Support_Public));

                _copyAndDelete = _moveGroups.Selected
                                        .Concat(_moveGroups.ExclusiveSupportDeclarations)
                                        .Except(publicSupportMembersToRetain)
                                        .Except(mustRetainPrivateDeclarations)
                                        .Except(privateDependenciesOfRetainedPublicSupportMembers)
                                        .ToList();
            }
            else
            {
                _copyAndDelete = _moveGroups.Selected
                        .Concat(_moveGroups.ExclusiveSupportDeclarations)
                        .ToList();

            }
            _retain = _moveGroups.AllParticipants.Except(_copyAndDelete).ToList();
        }

        private bool IsMustMovePublicSupport(IMoveableMemberSet publicSupportMember, IEnumerable<Declaration> mustMovePrivateSupport, out IEnumerable<Declaration> newMustMovePrivateSupport)
        {
            newMustMovePrivateSupport = new List<Declaration>();
            foreach (var mustMovePvtSupport in mustMovePrivateSupport)
            {
                if (publicSupportMember.DirectDependencies.Contains(mustMovePvtSupport))
                {
                    newMustMovePrivateSupport = publicSupportMember.DirectDependencies.Where(d => d.HasPrivateAccessibility());
                    return true;
                }
            }
            return false;
        }

        private bool HasUnMoveableSupportContent(IMoveGroupsProvider moveGroups, out IEnumerable<Declaration> unMoveables)
        {
            unMoveables = moveGroups.NonExclusiveSupportDeclarations.Where(d => d.HasPrivateAccessibility());
            return unMoveables.Any();
        }

        private bool PublicSupportMembersProvideUnMoveableDeclarationsAccess(IMoveGroupsProvider moveGroups)
        {
            var privateNonExclusiveReferences = moveGroups.NonExclusiveSupportDeclarations
                .Where(d => d.HasPrivateAccessibility())
                .AllReferences()
                .Where(rf => moveGroups.AllParticipants.Contains(rf.ParentScoping));

            var publicSupportingMembers = moveGroups.Declarations(MoveGroups.Support_Public);
            return publicSupportingMembers.ContainsParentScopesForAllReferences(privateNonExclusiveReferences);
        }

        private static IEnumerable<IGrouping<QualifiedModuleName, IdentifierReference>> RenameableReferencesByQualifiedModuleName(IEnumerable<IdentifierReference> references)
        {
            var modules = references
                .Where(reference =>
                    reference.Context.GetText() != Tokens.Me
                    && !reference.IsArrayAccess
                    && !reference.IsDefaultMemberAccess)
                .GroupBy(r => r.QualifiedModuleName);

            return modules;
        }

        private static IEnumerable<IdentifierReference> RenameableReferences(Declaration declaration, QualifiedModuleName qmn)
            => RenameableReferencesByQualifiedModuleName(declaration.References).Where(g => g.Key == qmn).SelectMany(g => g);


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
