using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.DeclareFieldsAsUDTMembers;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.CodeBlockInsert;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ICodeOnlyRefactoringAction<DeclareFieldsAsUDTMembersModel> _declareFieldAsUDTMemberRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replaceUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceFieldReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly ICodeBuilder _codeBuilder;
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replaceUDTMemberReferencesModelFactory;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IReplacePrivateUDTMemberReferencesModelFactory replaceUDTMemberReferencesModelFactory,
            IDeclarationFinderProvider declarationFinderProvider,
            IRewritingManager rewritingManager,
            ICodeBuilder codeBuilder)
                : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _declareFieldAsUDTMemberRefactoringAction = refactoringActionsProvider.DeclareFieldsAsUDTMembers;
            _replaceUDTMemberReferencesRefactoringAction = refactoringActionsProvider.ReplaceUDTMemberReferences;
            _replaceFieldReferencesRefactoringAction = refactoringActionsProvider.ReplaceReferences;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _codeBuilder = codeBuilder;
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

        private void ModifyFields(EncapsulateFieldModel encapsulateFieldModel, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(encapsulateFieldModel.QualifiedModuleName);

            rewriter.RemoveVariables(encapsulateFieldModel.SelectedFieldCandidates.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());

            if (encapsulateFieldModel.ObjectStateUDTField.IsExistingDeclaration)
            {
                var model = new DeclareFieldsAsUDTMembersModel();

                foreach (var field in encapsulateFieldModel.SelectedFieldCandidates)
                {
                    model.AssignFieldToUserDefinedType(encapsulateFieldModel.ObjectStateUDTField.AsTypeDeclaration, field.Declaration as VariableDeclaration, field.PropertyIdentifier);
                }
                _declareFieldAsUDTMemberRefactoringAction.Refactor(model, rewriteSession);
            }
            else
            {
                var newUDTMembers = encapsulateFieldModel.SelectedFieldCandidates
                    .Select(m => (m.Declaration as VariableDeclaration, m.BackingIdentifier));

                var typeDeclarationBlock =  _codeBuilder.BuildUserDefinedTypeDeclaration(encapsulateFieldModel.ObjectStateUDTField.AsTypeName, newUDTMembers);

                encapsulateFieldModel.AddContentBlock(NewContentType.TypeDeclarationBlock, typeDeclarationBlock);

                encapsulateFieldModel.AddContentBlock(NewContentType.DeclarationBlock, $"{Accessibility.Private} {encapsulateFieldModel.ObjectStateUDTField.IdentifierName} {Tokens.As} {encapsulateFieldModel.ObjectStateUDTField.AsTypeName}");
            }
        }

        private void ModifyReferences(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            var udtFields = model.SelectedFieldCandidates
                .Where(f => (f.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false)
                    && f.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private);

            if (udtFields.Any())
            {
                var replaceUDTMemberReferencesModel = _replaceUDTMemberReferencesModelFactory.Create(udtFields.Select(f => f.Declaration).Cast<VariableDeclaration>());

                foreach (var udtfield in udtFields)
                {
                    foreach (var udtMember in replaceUDTMemberReferencesModel.UDTMembers)
                    {
                        var localReplacement = udtfield.Declaration.IsArray 
                            ? $"{udtfield.IdentifierName}.{udtMember.IdentifierName.CapitalizeFirstLetter()}" 
                            : udtMember.IdentifierName.CapitalizeFirstLetter();

                        var udtExpressions = new PrivateUDTMemberReferenceReplacementExpressions($"{udtfield.IdentifierName}.{udtMember.IdentifierName}")
                        {
                            LocalReferenceExpression = udtMember.IdentifierName.CapitalizeFirstLetter(),
                        };

                        replaceUDTMemberReferencesModel.AssignUDTMemberReferenceExpressions(udtfield.Declaration as VariableDeclaration, udtMember, udtExpressions);
                    }
                    _replaceUDTMemberReferencesRefactoringAction.Refactor(replaceUDTMemberReferencesModel, rewriteSession);
                }
            }

            var modelReplaceField = new ReplaceReferencesModel()
            {
                ModuleQualifyExternalReferences = true,
            };

            foreach (var field in model.SelectedFieldCandidates.Except(udtFields))
            {
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementExpression = idRef.QualifiedModuleName == field.QualifiedModuleName
                        ? field.Declaration.IsArray ? $"{model.ObjectStateUDTField.FieldIdentifier}.{field.BackingIdentifier}" : field.PropertyIdentifier
                        : field.PropertyIdentifier;
                     modelReplaceField.AssignFieldReferenceReplacementExpression(idRef, replacementExpression);
                }

            }
            _replaceFieldReferencesRefactoringAction.Refactor(modelReplaceField, rewriteSession);
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
    }
}
