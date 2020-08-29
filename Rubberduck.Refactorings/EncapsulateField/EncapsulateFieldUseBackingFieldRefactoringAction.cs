using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.CodeBlockInsert;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldUseBackingFieldRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replaceUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceDeclarationIdentifierModel> _replaceDeclarationIdentifiers;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replaceUDTMemberReferencesModelFactory;

        public EncapsulateFieldUseBackingFieldRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IReplacePrivateUDTMemberReferencesModelFactory replaceUDTMemberReferencesModelFactory,
            IDeclarationFinderProvider declarationFinderProvider,
            IRewritingManager rewritingManager)
                :base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _replaceUDTMemberReferencesRefactoringAction = refactoringActionsProvider.ReplaceUDTMemberReferences;
            _replaceReferencesRefactoringAction = refactoringActionsProvider.ReplaceReferences;
            _replaceDeclarationIdentifiers = refactoringActionsProvider.ReplaceDeclarationIdentifiers;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _replaceUDTMemberReferencesModelFactory = replaceUDTMemberReferencesModelFactory;
        }

        public override void Refactor(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any())
            {
                return;
            }

            model.NewContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.CodeSectionBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };

            ModifyFields(model, rewriteSession);

            ModifyReferences(model, rewriteSession);

            InsertNewContent(model, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldModel model, IRewriteSession rewriteSession)
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

                model.AddContentBlock(NewContentType.DeclarationBlock, newField);
            }

            var retainedFieldDeclarations = model.SelectedFieldCandidates.Except(fieldDeclarationsToDeleteAndReplace).ToList();

            if (retainedFieldDeclarations.Any())
            {
                MakeImplicitDeclarationTypeExplicit(retainedFieldDeclarations, rewriter);

                SetPrivateVariableVisiblity(retainedFieldDeclarations, rewriter);

                Rename(retainedFieldDeclarations, rewriteSession);
            }
        }

        private void ModifyReferences(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            var privateUdtInstances = model.SelectedFieldCandidates
                .Where(f => (f.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false)
                    && f.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private);

            ReplaceEncapsulatedPrivateUserDefinedTypeMemberReferences(privateUdtInstances, rewriteSession);

            ReplaceEncapsulatedFieldReferences(model.SelectedFieldCandidates.Except(privateUdtInstances), rewriteSession);
        }

        private void InsertNewContent(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            var encapsulateFieldInsertNewCodeModel = new EncapsulateFieldInsertNewCodeModel(model.SelectedFieldCandidates)
            {
                NewContent = model.NewContent,
                IncludeNewContentMarker = model.IncludeNewContentMarker
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

                    model.AssignFieldReferenceReplacementExpression(idRef, replacementExpression);
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

        private static void MakeImplicitDeclarationTypeExplicit(IEnumerable<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
        {
            var fieldsToChange = fields.Where(f => !f.Declaration.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
                .Select(f => f.Declaration);

            foreach (var field in fieldsToChange)
            {
                rewriter.InsertAfter(field.Context.Stop.TokenIndex, $" {Tokens.As} {field.AsTypeName}");
            }
        }

        private static void SetPrivateVariableVisiblity(IEnumerable<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
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

        private void Rename(IEnumerable<IEncapsulateFieldCandidate> fields, IRewriteSession rewriteSession)
        {
            var fieldToNewNamePairs = fields.Where(f => !f.BackingIdentifier.Equals(f.Declaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase))
                .Select(f => (f.Declaration, f.BackingIdentifier));

            _replaceDeclarationIdentifiers.Refactor(new ReplaceDeclarationIdentifierModel(fieldToNewNamePairs), rewriteSession);
        }
    }
}
