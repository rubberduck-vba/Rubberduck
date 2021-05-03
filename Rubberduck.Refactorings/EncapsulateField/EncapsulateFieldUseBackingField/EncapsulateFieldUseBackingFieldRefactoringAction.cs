using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingField
{
    public class EncapsulateFieldUseBackingFieldRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldUseBackingFieldModel>
    {
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replaceUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> _replaceDeclarationIdentifiers;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replaceUDTMemberReferencesModelFactory;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;

        public EncapsulateFieldUseBackingFieldRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IReplacePrivateUDTMemberReferencesModelFactory replaceUDTMemberReferencesModelFactory,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory newContentAggregatorFactory)
                :base(rewritingManager)
        {
            _replaceUDTMemberReferencesRefactoringAction = refactoringActionsProvider.ReplaceUDTMemberReferences;
            _replaceReferencesRefactoringAction = refactoringActionsProvider.ReplaceReferences;
            _replaceDeclarationIdentifiers = refactoringActionsProvider.ReplaceDeclarationIdentifiers;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _replaceUDTMemberReferencesModelFactory = replaceUDTMemberReferencesModelFactory;
            _newContentAggregatorFactory = newContentAggregatorFactory;
        }

        public override void Refactor(EncapsulateFieldUseBackingFieldModel model, IRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any())
            {
                return;
            }

            var publicFieldsDeclaredInListsToReDeclareAsPrivateBackingFields 
                = model.SelectedFieldCandidates
                    .Where(f => f.Declaration.IsDeclaredInList()
                        && !f.Declaration.HasPrivateAccessibility())
                    .ToList();

            ModifyFields(model, publicFieldsDeclaredInListsToReDeclareAsPrivateBackingFields, rewriteSession);

            ModifyReferences(model, rewriteSession);

            InsertNewContent(model, publicFieldsDeclaredInListsToReDeclareAsPrivateBackingFields, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldUseBackingFieldModel model, List<IEncapsulateFieldCandidate> publicFieldsToRemove, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.QualifiedModuleName);
            rewriter.RemoveVariables(publicFieldsToRemove.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());

            var retainedFieldDeclarations = model.SelectedFieldCandidates
                .Except(publicFieldsToRemove)
                .ToList();

            if (retainedFieldDeclarations.Any())
            {
                MakeImplicitDeclarationTypeExplicit(retainedFieldDeclarations, rewriter);

                SetPrivateVariableVisiblity(retainedFieldDeclarations, rewriter);

                Rename(retainedFieldDeclarations, rewriteSession);
            }
        }

        private void ModifyReferences(EncapsulateFieldUseBackingFieldModel model, IRewriteSession rewriteSession)
        {
            var privateUdtInstances = model.SelectedFieldCandidates
                .Where(f => (f.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false)
                    && f.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private);

            ReplaceUDTMemberReferencesOfPrivateUDTFields(privateUdtInstances, rewriteSession);

            ReplaceEncapsulatedFieldReferences(model.SelectedFieldCandidates.Except(privateUdtInstances), rewriteSession);
        }

        private void InsertNewContent(EncapsulateFieldUseBackingFieldModel model, List<IEncapsulateFieldCandidate> candidatesRequiringNewBackingFields, IRewriteSession rewriteSession)
        {
            var aggregator = model.NewContentAggregator ?? _newContentAggregatorFactory.Create();
            model.NewContentAggregator = null;

            var encapsulateFieldInsertNewCodeModel = new EncapsulateFieldInsertNewCodeModel(model.SelectedFieldCandidates)
            {
                CandidatesRequiringNewBackingFields = candidatesRequiringNewBackingFields,
                NewContentAggregator = aggregator
            };

            _encapsulateFieldInsertNewCodeRefactoringAction.Refactor(encapsulateFieldInsertNewCodeModel, rewriteSession);
        }

        private void ReplaceEncapsulatedFieldReferences(IEnumerable<IEncapsulateFieldCandidate> fieldCandidates, IRewriteSession rewriteSession)
        {
            var model = new ReplaceReferencesModel()
            {
                ModuleQualifyExternalReferences = true
            };

            foreach (var field in fieldCandidates)
            {
                InitializeModel(model, field);
            }

            _replaceReferencesRefactoringAction.Refactor(model, rewriteSession);
        }

        private void ReplaceUDTMemberReferencesOfPrivateUDTFields(IEnumerable<IEncapsulateFieldCandidate> udtFieldCandidates, IRewriteSession rewriteSession)
        {
            if (!udtFieldCandidates.Any())
            {
                return;
            }

            var replacePrivateUDTMemberReferencesModel 
                = _replaceUDTMemberReferencesModelFactory.Create(udtFieldCandidates.Select(f => f.Declaration).Cast<VariableDeclaration>());

            foreach (var udtfield in udtFieldCandidates)
            {
                InitializeModel(replacePrivateUDTMemberReferencesModel, udtfield);
            }
            _replaceUDTMemberReferencesRefactoringAction.Refactor(replacePrivateUDTMemberReferencesModel, rewriteSession);
        }

        private void InitializeModel(ReplaceReferencesModel model, IEncapsulateFieldCandidate field)
        {
            foreach (var idRef in field.Declaration.References)
            {
                var replacementExpression = field.PropertyIdentifier;

                if (idRef.QualifiedModuleName == field.QualifiedModuleName && field.Declaration.IsArray)
                {
                    replacementExpression = field.BackingIdentifier;
                }

                model.AssignReferenceReplacementExpression(idRef, replacementExpression);
            }
        }

        private void InitializeModel(ReplacePrivateUDTMemberReferencesModel model, IEncapsulateFieldCandidate udtfield)
        {
            foreach (var udtMember in model.UDTMembers)
            {
                var udtExpressions = new PrivateUDTMemberReferenceReplacementExpressions($"{udtfield.IdentifierName}.{udtMember.IdentifierName}")
                {
                    LocalReferenceExpression = udtMember.IdentifierName,
                };

                model.AssignUDTMemberReferenceExpressions(udtfield.Declaration as VariableDeclaration, udtMember, udtExpressions);
            }
        }

        private static void MakeImplicitDeclarationTypeExplicit(IReadOnlyCollection<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
        {
            var fieldsToChange = fields.Where(f => !f.Declaration.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
                .Select(f => f.Declaration);

            foreach (var field in fieldsToChange)
            {
                rewriter.InsertAfter(field.Context.Stop.TokenIndex, $" {Tokens.As} {field.AsTypeName}");
            }
        }

        private static void SetPrivateVariableVisiblity(IReadOnlyCollection<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
        {
            var visibility = Accessibility.Private.TokenString();
            foreach (var element in fields.Where(f => !f.Declaration.HasPrivateAccessibility()).Select(f => f.Declaration))
            {
                if (!element.IsVariable())
                {
                    throw new ArgumentException();
                }

                var variableStmtContext = element.Context.GetAncestor<VBAParser.VariableStmtContext>();
                var visibilityContext = variableStmtContext.GetChild<VBAParser.VisibilityContext>();

                if (visibilityContext != null)
                {
                    rewriter.Replace(visibilityContext, visibility);
                    continue;
                }
                rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{visibility} ");
            }
        }

        private void Rename(IReadOnlyCollection<IEncapsulateFieldCandidate> fields, IRewriteSession rewriteSession)
        {
            var fieldToNewNamePairs = fields.Where(f => !f.BackingIdentifier.Equals(f.Declaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase))
                .Select(f => (f.Declaration, f.BackingIdentifier));

            var model = new ReplaceDeclarationIdentifierModel(fieldToNewNamePairs);
            _replaceDeclarationIdentifiers.Refactor(model, rewriteSession);
        }
    }
}
