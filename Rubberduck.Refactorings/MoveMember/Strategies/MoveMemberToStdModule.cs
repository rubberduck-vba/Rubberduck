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
    public enum MoveDisposition { Retain, Move }

    public interface IMoveMemberRefactoringStrategy
    {
        IMoveMemberRewriteSession ModifyContent(MoveMemberModel model, IExecutableRewriteSession session, IProvideNewContent contentToMove);
    }

    public abstract class MoveMemberRefactoringStrategyBase
    {
        public abstract bool IsApplicable(MoveMemberModel model);
    }

    public class MoveMemberToStdModule : MoveMemberRefactoringStrategyBase, IMoveMemberRefactoringStrategy
    {
        public override bool IsApplicable(MoveMemberModel model)
        {
            if (!model.IsStdModuleDestination) { return false; }

            //Strategy does not move fields of Private UserDefinedTypes
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable) 
                && (nm.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
                && nm.AsTypeDeclaration.HasPrivateAccessibility())) { return false; }

            //Strategy does not move fields of Private EnumTypes
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                && (nm.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.Enumeration) ?? false)
                && nm.AsTypeDeclaration.HasPrivateAccessibility())) { return false; }

            var moveGroups = model.MoveGroups;

            if (!model.IsStdModuleSource)
            {
                //External references to public elements 
                //of Classes and Forms are not supported for moving in this strategy
                var externalMemberRefs = model.SelectedDeclarations.AllReferences().Where(rf => rf.QualifiedModuleName != model.Source.QualifiedModuleName);

                if (externalMemberRefs.Any())
                {
                    return false;
                }

                return !moveGroups.AllNonExclusiveSupportDeclarations.Any();
            }

            var unmoveableDeclarations = moveGroups.AllNonExclusiveSupportDeclarations.ToList();
            if (!unmoveableDeclarations.Any(ud => ud.HasPrivateAccessibility()))
            {
                return true;
            }

            var unmoveableMembers = unmoveableDeclarations.Where(umd => umd.IsMember() && umd.HasPrivateAccessibility());
            if (unmoveableMembers.Any())
            {
                return false;
            }

            var callBackMembers = unmoveableDeclarations.Except(unmoveableMembers).Where(um => um.IsMember());
            if (callBackMembers.Any())
            {
                var callBackCallChainElements = moveGroups.CallChainDeclarations(callBackMembers) //, scenario.DeclarationFinderProvider)
                    .Except(callBackMembers);

                var exclusiveCallBackDeclarations = callBackCallChainElements
                    .Where(supporting => supporting.References.All(seRefs => callBackMembers.Contains(seRefs.ParentScoping)));

                var unmoveableNonMembers = unmoveableDeclarations.Where(umd => !umd.IsMember());
                if (unmoveableNonMembers.All(umnm => exclusiveCallBackDeclarations.Contains(umnm)))
                {
                    return true;
                }
            }
            return false;
        }

        private MoveMemberModel _model;
        private IDeclarationFinderProvider _declarationFinderProvider;
        private IProvideNewContent _contentToMove;

        public IMoveMemberRewriteSession ModifyContent(MoveMemberModel model, IExecutableRewriteSession executableRewriteSession, IProvideNewContent contentToMove)
        {

            _declarationFinderProvider = model.DeclarationFinderProvider;
            _model = model;
            _contentToMove = contentToMove;

            var rewriteSession = new MoveMemberRewriteSession(executableRewriteSession);
#if DEBUG
            var preModify = GetContentState();
#endif
            rewriteSession.Remove(this[MoveDisposition.Move]);

            ModifyRetainedSourceContent(rewriteSession);

#if DEBUG
            var postSourceModify = GetContentState();
#endif
            if (model.Destination.IsExistingModule(out var destination))
            {
                ModifyExistingDestinationContentAffectedByMove(destination, rewriteSession);

                InsertMovedContent(rewriteSession, destination, MovedContent);
            }
#if DEBUG
            var postDestinationModify = GetContentState();
#endif
            UpdateCallSites(rewriteSession);
#if DEBUG
            var finalState = GetContentState();
#endif
            return rewriteSession;
        }

        private string MovedContent
        {
            get
            {
                var scratchPadRewriteSession = _model.MoveRewritingManager.CheckOutCodePaneSession();
                var movedContentRewriter = scratchPadRewriteSession.CheckOutModuleRewriter(_model.Source.QualifiedModuleName);

                foreach (var element in this[MoveDisposition.Move]/*.Where(m => m.IsMember())*/)
                {
                    var codeBlock = DestinationMemberCodeBlock(element, movedContentRewriter);
                    if (element.IsMember())
                    {
                        _contentToMove.AddMethod(codeBlock);
                    }
                    else
                    {
                        _contentToMove.AddFieldOrConstantDeclaration(codeBlock);
                    }
                }

                return _contentToMove.AsSingleBlock;
            }
        }

        public string PreviewDestination(MoveMemberModel model)
        {
            var preview = string.Empty;
            if (model.Destination.IsExistingModule(out var module))
            {
                var session = new MoveMemberRewriteSession(model.MoveRewritingManager.CheckOutCodePaneSession());
                var rewriter = InsertMovedContent(session, module, MovedContent);
                preview = rewriter.GetText();
            }
            else
            {
                preview = $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}{MovedContent}";
            }

            return preview;
        }

        private IEnumerable<Declaration> this[MoveDisposition disposition]
        {
            get
            {
                switch (disposition)
                {
                    case MoveDisposition.Move:
                        return _model.MoveGroups.Selected.Concat(_model.MoveGroups.AllExclusiveSupportDeclarations);
                    case MoveDisposition.Retain:
                        return _model.MoveGroups.AllParticipants.Except(_model.MoveGroups.Selected.Concat(_model.MoveGroups.AllExclusiveSupportDeclarations));
                }
                return Enumerable.Empty<Declaration>();
            }
        }

        private void ModifyRetainedSourceContent(IMoveMemberRewriteSession rewriteSession)
        {
            var movedDeclarationIdentifierReferencesToModuleQualify = RenameSupport.RenameReferencesByQualifiedModuleName(_model.MoveGroups.Selected.AllReferences())
                .Where(g => g.Key == _model.Source.QualifiedModuleName)
                .SelectMany(grp => grp)
                .Where(rf => !this[MoveDisposition.Move].Contains(rf.ParentScoping));

            if (movedDeclarationIdentifierReferencesToModuleQualify.Any())
            {
                var sourceRewriter = rewriteSession.CheckOutModuleRewriter(_model.Source.QualifiedModuleName);
                foreach (var rf in movedDeclarationIdentifierReferencesToModuleQualify)
                {
                    sourceRewriter.Replace(rf.Context, AddDestinationModuleQualification(rf));
                }
            }
        }

        private void ModifyExistingDestinationContentAffectedByMove(Declaration destination, IMoveMemberRewriteSession rewriteSession)
        {
            var destinationReferencesToMovedMembers = _model.MoveGroups.Selected.AllReferences()
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
            if (_model.Destination.TryGetCodeSectionStartIndex(_declarationFinderProvider, out var codeSectionStartIndex))
            {
                destinationRewriter.InsertBefore(codeSectionStartIndex, $"{movedContent}{Environment.NewLine}{Environment.NewLine}");
            }
            else
            {
                destinationRewriter.InsertAtEndOfFile($"{Environment.NewLine}{Environment.NewLine}{movedContent}");
            }

            return destinationRewriter;
        }

        private string DestinationMemberCodeBlock(Declaration member, IModuleRewriter rewriter)
        {
            if (member.IsVariable() || member.IsConstant())
            {
                var visibility = IsOnlyReferencedByMovedMethods(member) ? Tokens.Private : Tokens.Public;

                return member.IsVariable()
                    ? $"{visibility} {rewriter.GetText(member)}"
                    : $"{visibility} {Tokens.Const} {rewriter.GetText(member)}"; //.Context.Start.TokenIndex, member.Context.Stop.TokenIndex)}";
            }

            if (_model.MoveGroups.Selected.Contains(member)
                || !IsOnlyReferencedByMovedMethods(member))
            {
                rewriter.SetMemberAccessibility(member, Tokens.Public);
            }

            if (_model.IsStdModuleSource)
            {
                rewriter = AddSourceModuleQualificationToMovedReferences(member, rewriter);
            }

            var destinationMemberAccessReferencesToMovedMembers = _model.AllDestinationModuleDeclarations
                .AllReferences().Where(rf => rf.ParentScoping == member);

            rewriter.RemoveMemberAccess(destinationMemberAccessReferencesToMovedMembers);
            return rewriter.GetText(member);
        }

        private IModuleRewriter AddSourceModuleQualificationToMovedReferences(Declaration member, IModuleRewriter tempRewriter)
        {
            var retainedPublicDeclarations = this[MoveDisposition.Retain].Where(m => !m.HasPrivateAccessibility());
            if (retainedPublicDeclarations.Any())
            {
                var destinationRefs = retainedPublicDeclarations.AllReferences().Where(rf => this[MoveDisposition.Move].Contains(rf.ParentScoping));
                foreach (var rf in destinationRefs)
                {
                    tempRewriter.Replace(rf.Context, $"{_model.Source.ModuleName}.{rf.IdentifierName}");
                }
            }
            return tempRewriter;
        }

        private void UpdateCallSites(IMoveMemberRewriteSession rewriteSession)
        {
            if (!_model.IsStdModuleDestination)
            {
                return;
            }

            var qmnToReferenceIdentifiers = RenameSupport.RenameReferencesByQualifiedModuleName(_model.MoveGroups.Selected.AllReferences());

            foreach (var references in qmnToReferenceIdentifiers.Where(qmn => !IsMoveEndpoint(qmn.Key)))
            {
                var moduleRewriter = rewriteSession.CheckOutModuleRewriter(references.Key);

                var memberAccessExprContexts = references.Where(rf => rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)
                        .Select(rf => rf.Context.Parent as VBAParser.MemberAccessExprContext);

                var destinationModuleName = _model.Destination.ModuleName;
                foreach (var memberAccessExprContext in memberAccessExprContexts)
                {
                    moduleRewriter.Replace(memberAccessExprContext.lExpression(), destinationModuleName);
                }

                var withMemberAccessExprContexts = references.Where(rf => rf.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        .Select(rf => rf.Context.Parent as VBAParser.WithMemberAccessExprContext);

                foreach (var wma in withMemberAccessExprContexts)
                {
                    moduleRewriter.InsertBefore(wma.Start.TokenIndex, destinationModuleName);
                }

                var nonQualifiedReferences = references.Where(rf => !(rf.Context.Parent is VBAParser.WithMemberAccessExprContext
                    || (rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)));

                foreach (var rf in nonQualifiedReferences)
                {
                    moduleRewriter.Replace(rf.Context, $"{destinationModuleName}.{rf.Declaration.IdentifierName}");
                }
            }
        }

        private bool IsMoveEndpoint(QualifiedModuleName qmn)
        {
            return qmn == _model.Source.QualifiedModuleName
                || (_model.Destination.IsExistingModule(out var module) && qmn == module.QualifiedModuleName);
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

            if (this[MoveDisposition.Retain].Contains(identifierReference.Declaration))
            {
                return identifierReference.IdentifierName;
            }
            return $"{_model.Destination.ModuleName}.{identifierReference.IdentifierName}";
        }

        private bool IsOnlyReferencedByMovedMethods(Declaration element)
            => element.References.All(rf => this[MoveDisposition.Move].Where(m => m.IsMember()).Contains(rf.ParentScoping));

        private (string Source, string Destination) GetContentState()
        {
#if DEBUG
            var tempSession = _model.MoveRewritingManager.CheckOutCodePaneSession();
            var sourceRewriter = tempSession.CheckOutModuleRewriter(_model.Source.QualifiedModuleName);
            if (_model.Destination.IsExistingModule(out var module))
            {
                var destinationRewriter = tempSession.CheckOutModuleRewriter(module.QualifiedModuleName);

                var newSource = sourceRewriter.GetText();
                var newDestination = destinationRewriter.GetText();
                return (newSource, newDestination);
            }
#endif
            return (string.Empty, string.Empty);
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
