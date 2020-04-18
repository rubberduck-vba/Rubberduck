using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Refactorings.Rename;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberStdToClassModule : MoveMemberStrategyBase, IMoveMemberRefactoringStrategy
    {
        public MoveMemberStdToClassModule(IDeclarationFinderProvider declarationFinderProvider,
                                        RenameCodeDefinedIdentifierRefactoringAction renameAction,
                                        IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory,
                                        INameConflictFinder nameConflictFinder,
                                        IDeclarationProxyFactory declarationProxyFactory)
            : base(declarationFinderProvider,
                                        renameAction,
                                        moveGroupsProviderFactory,
                                        nameConflictFinder,
                                        declarationProxyFactory)
        {}

        public override bool IsApplicable(MoveMemberModel model)
        {
            //TODO: Implement StdToClass moves
            return false;

            if (!model.Destination.IsClassModule) { return false; }

            var moveGroups = _moveGroupsProviderFactory.Create(model.MoveableMemberSets);

            //Strategy does not support Private Fields as a SelectedDeclaration.  But, the strategy
            //is still applicable if ALL of the Selected Fields are contained in the MoveGroup.Support_Exclusive.
            //Private/MoveGroup.Support_Exclusive are always moved along with the moved Members that reference/depend on them
            if (moveGroups.SelectedPrivateFields.Except(moveGroups.Declarations(MoveGroup.Support_Exclusive)).Any())
            {
                return false;
            }

            return TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out _);
        }

        public bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage)
        {
            //TODO: Implement StdToClass moves
            nonExecutableMessage = string.Empty;
            return false;

            return base.IsExecutableModelBase(model, out nonExecutableMessage);
        }

        protected override Dictionary<MoveDisposition, List<Declaration>> DetermineDispositionGroups(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups)
        {
            var dispositions = EmptyDispositions();
            if (model.Source.IsStandardModule
                && TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out dispositions))
            {
                return dispositions;
            }

            throw new MoveMemberUnsupportedMoveException();
        }

        protected override void ModifyRetainedReferencesToMovedMembers(IRewriteSession rewriteSession, MoveMemberModel model, Dictionary<MoveDisposition, List<Declaration>> dispositions)
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

        protected override void InsertNewSourceContent(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, IRewriteSession moveMemberRewriteSession, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions, IMovedContentProvider sourceContentProvider)
        {
            sourceContentProvider = LoadSourceNewContentProvider(sourceContentProvider, model, moveGroups, scratchPadSession, dispositions);

            var newContent = sourceContentProvider.AsSingleBlock;

            InsertNewContent(moveMemberRewriteSession, model.Source, newContent);
        }

        protected IMovedContentProvider LoadSourceNewContentProvider(IMovedContentProvider contentProvider, MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, IRewriteSession scratchPadSession, Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            var instanceFieldIdentiier = CreatNonConflictObjectFieldIdentifier(model.Source.Module as ModuleDeclaration, model.Destination.ModuleName);

            var destinationClassDeclaration = $"{Tokens.Private} {instanceFieldIdentiier} {Tokens.As} {model.Destination.ModuleName}";

            var destinationClassInstantiation = $"{Tokens.Set} {instanceFieldIdentiier} = {Tokens.New} {model.Destination.ModuleName}";

            var indent4Spaces = "    ";
            var pvtPropertySignature = $"{Tokens.Private} {Tokens.Property} {Tokens.Get} {model.Destination.ModuleName}() {Tokens.As} {model.Destination.ModuleName}";
            var pvtPropertyBodyIf = $"{indent4Spaces}{Tokens.If} {instanceFieldIdentiier} {Tokens.Is} {Tokens.Nothing} {Tokens.Then}";
            var pvtPropertyBodyInstantiation = $"{indent4Spaces}{indent4Spaces}{destinationClassInstantiation}";
            var pvtPropertyBodyEndIf = $"{indent4Spaces}{Tokens.End} {Tokens.If}";
            var pvtPropertyAssignment = $"{indent4Spaces}{Tokens.Set} {model.Destination.ModuleName} = {instanceFieldIdentiier}";
            var pvtPropertyEnd = $"{Tokens.End} {Tokens.Property}";

            var pvtProperty = string.Join(Environment.NewLine, pvtPropertySignature,
                                                                pvtPropertyBodyIf,
                                                                pvtPropertyBodyInstantiation,
                                                                pvtPropertyBodyEndIf,
                                                                pvtPropertyAssignment,
                                                                pvtPropertyEnd);


            contentProvider.AddFieldOrConstantDeclaration(destinationClassDeclaration);
            contentProvider.AddMethodDeclaration(pvtProperty);

            return contentProvider;
        }

        private string CreatNonConflictObjectFieldIdentifier(ModuleDeclaration sourceModule, string destinationModuleName)
        {
            var fieldIdentifier = $"{destinationModuleName.ToLowerCaseFirstLetter().IncrementIdentifier()}";

            var members = _declarationFinderProvider.DeclarationFinder.Members(sourceModule.QualifiedModuleName)
                .Where(d => d.IsMemberVariable() && destinationModuleName.IsEquivalentVBAIdentifierTo(d.AsTypeDeclaration.IdentifierName));

            if (members.Any() && members.Count() == 1)
            {
                return members.Single().IdentifierName;
            }

            var instanceProxy = _declarationProxyFactory.Create(fieldIdentifier, DeclarationType.Variable, Accessibility.Private, sourceModule as ModuleDeclaration, sourceModule);

            var conflictsResolved = false;
            for (var idx = 0; idx < 20 && !conflictsResolved; idx++)
            {
                if (!_nameConflictFinder.ProposedDeclarationCreatesConflict(instanceProxy, out _))
                {
                    conflictsResolved = true;
                    continue;
                }
                instanceProxy.IdentifierName = instanceProxy.IdentifierName.IncrementIdentifier();
            }
            return instanceProxy.IdentifierName;
        }
    }
}
