using Rubberduck.Common;
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

            if (model.NewContentAggregator is null)
            {
                model.NewContentAggregator = _newContentAggregatorFactory.Create();
            }

            ModifyFields(model, rewriteSession);

            ModifyReferences(model, rewriteSession);

            InsertNewContent(model, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldUseBackingFieldModel model, IRewriteSession rewriteSession)
        {
            var fieldDeclarationsToDeleteAndReplace = model.SelectedFieldCandidates
                .Where(f => f.Declaration.IsDeclaredInList() 
                    && !f.Declaration.HasPrivateAccessibility())
                .ToList();

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.QualifiedModuleName);
            rewriter.RemoveVariables(fieldDeclarationsToDeleteAndReplace.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());

            foreach (var field in fieldDeclarationsToDeleteAndReplace)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.BackingIdentifier);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                model.NewContentAggregator.AddNewContent(NewContentType.DeclarationBlock, newField);
            }

            var retainedFieldDeclarations = model.SelectedFieldCandidates
                .Except(fieldDeclarationsToDeleteAndReplace)
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

            ReplaceEncapsulatedPrivateUserDefinedTypeMemberReferences(privateUdtInstances, rewriteSession);

            ReplaceEncapsulatedFieldReferences(model.SelectedFieldCandidates.Except(privateUdtInstances), rewriteSession);
        }

        private void InsertNewContent(EncapsulateFieldUseBackingFieldModel model, IRewriteSession rewriteSession)
        {
            var encapsulateFieldInsertNewCodeModel = new EncapsulateFieldInsertNewCodeModel(model.SelectedFieldCandidates)
            {
                NewContentAggregator = model.NewContentAggregator
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
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementExpression = idRef.QualifiedModuleName == field.QualifiedModuleName
                        ? field.Declaration.IsArray ? field.BackingIdentifier : field.PropertyIdentifier
                        : field.PropertyIdentifier;

                    model.AssignReferenceReplacementExpression(idRef, replacementExpression);
                }
            }

            _replaceReferencesRefactoringAction.Refactor(model, rewriteSession);
        }

        private void ReplaceEncapsulatedPrivateUserDefinedTypeMemberReferences(IEnumerable<IEncapsulateFieldCandidate> udtFieldCandidates, IRewriteSession rewriteSession)
        {
            if (!udtFieldCandidates.Any())
            {
                return;
            }

            var replacePrivateUDTMemberReferencesModel = _replaceUDTMemberReferencesModelFactory.Create(udtFieldCandidates.Select(f => f.Declaration).Cast<VariableDeclaration>());

            foreach (var udtfield in udtFieldCandidates)
            {
                foreach (var udtMember in replacePrivateUDTMemberReferencesModel.UDTMembers)
                {
                    var udtExpressions = new PrivateUDTMemberReferenceReplacementExpressions($"{udtfield.IdentifierName}.{udtMember.IdentifierName}")
                    {
                        LocalReferenceExpression = udtMember.IdentifierName.CapitalizeFirstLetter(),
                    };

                    replacePrivateUDTMemberReferencesModel.AssignUDTMemberReferenceExpressions(udtfield.Declaration as VariableDeclaration, udtMember, udtExpressions);
                }
                _replaceUDTMemberReferencesRefactoringAction.Refactor(replacePrivateUDTMemberReferencesModel, rewriteSession);
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
