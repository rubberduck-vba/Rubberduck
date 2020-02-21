using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToStdModule : IMoveMemberRefactoringStrategy
    {
        private MoveMemberModel _model;
        private IRewritingManager _rewritingManager;
        private IMoveGroupsProvider _moveGroups;
        private List<Declaration> _copyAndDelete;
        private List<Declaration> _retain;

        public bool IsApplicable(MoveMemberModel model)
        {
            //A strategy exists to handle 'nothing selected'...it's not this one
            if (!model.SelectedDeclarations.Any()) { return false; }

            if (!model.Destination.IsStandardModule) { return false; }

            //Strategy does not move fields of Private UserDefinedTypes
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable) 
                && (nm.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                && nm.AsTypeDeclaration.HasPrivateAccessibility())) { return false; }

            //Strategy does not move fields of Private EnumTypes
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                && (nm.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.Enumeration) ?? false)
                && nm.AsTypeDeclaration.HasPrivateAccessibility())) { return false; }

            //Strategy does not move fields of ObjectTypes - would disconnect declaration from instantiation execution paths
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && nm.IsObject)) { return false; }

            var moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.SelectedDeclarations);


            if (!model.Source.IsStandardModule)
            {
                if (moveGroups.NonExclusiveSupportDeclarations.Any()) { return false; }

                /*
                 * If there are external references to the selected element(s) and the source is a 
                 * ClassModule or  UserForm, then they cannot be supported by this strategy
                */
                var externalMemberRefs = model.SelectedDeclarations.AllReferences().Where(rf => rf.QualifiedModuleName != model.Source.QualifiedModuleName);
                return !externalMemberRefs.Any();
            }

            if (!HasUnMoveableSupportContent(moveGroups)) { return true; }

            return TryVerifySupportMembersProvideUnMoveableDeclarationsAccess(moveGroups, out _);
        }

        public bool IsAnExecutableStrategy => true;

        public IMoveMemberRewriteSession RefactorRewrite(MoveMemberModel model, IRewritingManager rewritingManager, INewContentProvider contentToMove, bool asPreview = false)
        {
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

        private void ModifyMovedIdentifiersToAvoidDestinationNameConflicts(MoveMemberModel model, IEnumerable<Declaration> copyAndDelete)
        {
            if (!model.Destination.IsExistingModule(out var destination)) { return; }

            var allPotentialConflictNames = model.Destination.ModuleDeclarations
                .Where(d => !NeverCauseNameConflictTypes.Contains(d.DeclarationType))
                .Select(m => m.IdentifierName);

            var conflictingIdentifiers = copyAndDelete
                        .Where(m => allPotentialConflictNames.Contains(m.IdentifierName))
                        .Select(m => m.IdentifierName);

            var identifier = string.Empty;
            var guard = 0;
            var maxIterations = 50;
            foreach (var conflictIdentifier in conflictingIdentifiers)
            {
                var moveable = model.MoveableMemberSetByName(conflictIdentifier);
                identifier = conflictIdentifier;
                guard = 0;
                while (guard++ < maxIterations && allPotentialConflictNames.Contains(identifier))
                {
                    identifier = identifier.IncrementIdentifier();
                }

                if (guard >= maxIterations)
                {
                    //If we end up here, something is probably really wrong and we should not proceed with the refactoring
                    throw new MoveMemberUnsupportedMoveException(moveable.Member);
                }

                moveable.MovedIdentifierName = identifier;
            }
        }

        private void ModifyRetainedReferencesToMovedMembers(IMoveMemberRewriteSession rewriteSession)
        {
            var retainedReferencesToModuleQualify = RenameableReferencesByQualifiedModuleName(_moveGroups.CallTreeRoots.AllReferences())
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
            var destinationReferencesToMovedMembers = _moveGroups.CallTreeRoots.AllReferences()
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
            var isOnlyReferencedByMovedMethods = member
                        .References.All(rf => _copyAndDelete
                            .Where(m => m.IsMember()).Contains(rf.ParentScoping));

            var membersRelatedByName = model.MoveableMemberSetByName(member.IdentifierName);

            if (member.IsVariable() || member.IsConstant())
            {
                var visibility  = isOnlyReferencedByMovedMethods  ? Tokens.Private : Tokens.Public;

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

            if (_moveGroups.CallTreeRoots.Contains(member)
                || !isOnlyReferencedByMovedMethods)
            {
                rewriter.SetMemberAccessibility(member, Tokens.Public);
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
                    = RenameableReferencesByQualifiedModuleName(_moveGroups.CallTreeRoots.AllReferences())
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

            _moveGroups = model.MoveMemberFactory.CreateMoveGroupsProvider(model.SelectedDeclarations);

            _copyAndDelete = _moveGroups.CallTreeRoots
                                        .Concat(_moveGroups.ExclusiveSupportDeclarations)
                                        .ToList();

            if (HasUnMoveableSupportContent(_moveGroups))
            {
                if (TryVerifySupportMembersProvideUnMoveableDeclarationsAccess(_moveGroups, out var supportMembers))
                {
                    //If we retain any public property member, then retain every L\S\G method for the identifier
                    var publicSupportMembersToRetain = supportMembers.Select(p => model.MoveableMemberSetByName(p.IdentifierName))
                                                            .SelectMany(mbn => mbn.Members).Distinct();

                    _copyAndDelete = _copyAndDelete.Except(publicSupportMembersToRetain).ToList();
                }
            }

            _retain = _moveGroups.AllParticipants.Except(_copyAndDelete).ToList();
        }

        private bool HasUnMoveableSupportContent(IMoveGroupsProvider moveGroups) 
                => moveGroups.NonExclusiveSupportDeclarations.Where(d => d.HasPrivateAccessibility()).Any();

        private bool TryVerifySupportMembersProvideUnMoveableDeclarationsAccess(IMoveGroupsProvider moveGroups, out List<Declaration> supportMembersToRetain)
        {
            var theReferencesThatMatter = moveGroups.NonExclusiveSupportDeclarations
                .Where(d => d.HasPrivateAccessibility())
                .AllReferences()
                .Where(rf => moveGroups.AllParticipants.Contains(rf.ParentScoping));

            var publicSupportingMembers = moveGroups.SupportMembers.Where(sm => !sm.HasPrivateAccessibility());

            supportMembersToRetain = new List<Declaration>();
            if(publicSupportingMembers.ContainsParentScopesForAllReferences(theReferencesThatMatter))
            {
                supportMembersToRetain = theReferencesThatMatter.Select(rf => rf.ParentScoping).Distinct().ToList();
                return true;
            }
            return false;
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
