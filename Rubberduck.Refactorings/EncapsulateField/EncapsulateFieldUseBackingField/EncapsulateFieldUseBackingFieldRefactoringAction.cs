using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using Rubberduck.Refactorings.DeleteDeclarations;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingField
{
    public class EncapsulateFieldUseBackingFieldRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldUseBackingFieldModel>
    {
        private readonly ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> _replaceDeclarationIdentifiers;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<DeleteDeclarationsModel> _deleteDeclarationsRefactoringAction;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;
        private readonly IEncapsulateFieldReferenceReplacerFactory _encapsulateFieldReferenceReplacerFactory;

        public EncapsulateFieldUseBackingFieldRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IEncapsulateFieldReferenceReplacerFactory encapsulateFieldReferenceReplacerFactory,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory newContentAggregatorFactory)
                :base(rewritingManager)
        {
            _replaceDeclarationIdentifiers = refactoringActionsProvider.ReplaceDeclarationIdentifiers;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _deleteDeclarationsRefactoringAction = refactoringActionsProvider.DeleteDeclarations;
            _newContentAggregatorFactory = newContentAggregatorFactory;
            _encapsulateFieldReferenceReplacerFactory = encapsulateFieldReferenceReplacerFactory;
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

            var referenceReplacer = _encapsulateFieldReferenceReplacerFactory.Create();
            referenceReplacer.ReplaceReferences(model.SelectedFieldCandidates, rewriteSession);

            InsertNewContent(model, publicFieldsDeclaredInListsToReDeclareAsPrivateBackingFields, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldUseBackingFieldModel model, List<IEncapsulateFieldCandidate> publicFieldsToRemove, IRewriteSession rewriteSession)
        {
            var deletionsModel = new DeleteDeclarationsModel(publicFieldsToRemove.Select(f => f.Declaration));
            
            _deleteDeclarationsRefactoringAction.Refactor(deletionsModel, rewriteSession);

            var retainedFieldDeclarations = model.SelectedFieldCandidates
                .Except(publicFieldsToRemove)
                .ToList();

            if (retainedFieldDeclarations.Any())
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(model.QualifiedModuleName);

                MakeImplicitDeclarationTypeExplicit(retainedFieldDeclarations, rewriter);

                SetPrivateVariableVisiblity(retainedFieldDeclarations, rewriter);

                Rename(retainedFieldDeclarations, rewriteSession);
            }
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
