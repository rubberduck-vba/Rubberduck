using Rubberduck.Common;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.DeclareFieldsAsUDTMembers;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldUseBackingUDTMemberModel>
    {
        private readonly ICodeOnlyRefactoringAction<DeclareFieldsAsUDTMembersModel> _declareFieldAsUDTMemberRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replaceUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceFieldReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly IEncapsulateFieldCodeBuilder _encapsulateFieldCodeBuilder;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replaceUDTMemberReferencesModelFactory;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IReplacePrivateUDTMemberReferencesModelFactory replaceUDTMemberReferencesModelFactory,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory newContentAggregatorFactory,
            IEncapsulateFieldCodeBuilderFactory encapsulateFieldCodeBuilderFactory)
                : base(rewritingManager)
        {
            _declareFieldAsUDTMemberRefactoringAction = refactoringActionsProvider.DeclareFieldsAsUDTMembers;
            _replaceUDTMemberReferencesRefactoringAction = refactoringActionsProvider.ReplaceUDTMemberReferences;
            _replaceFieldReferencesRefactoringAction = refactoringActionsProvider.ReplaceReferences;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _encapsulateFieldCodeBuilder = encapsulateFieldCodeBuilderFactory.Create();
            _replaceUDTMemberReferencesModelFactory = replaceUDTMemberReferencesModelFactory;
            _newContentAggregatorFactory = newContentAggregatorFactory;
        }

        public override void Refactor(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
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

        private void ModifyFields(EncapsulateFieldUseBackingUDTMemberModel encapsulateFieldModel, IRewriteSession rewriteSession)
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
                var objectStateTypeDeclarationBlock = _encapsulateFieldCodeBuilder.BuildUserDefinedTypeDeclaration(encapsulateFieldModel.ObjectStateUDTField, encapsulateFieldModel.EncapsulationCandidates);

                encapsulateFieldModel.NewContentAggregator.AddNewContent(NewContentType.UserDefinedTypeDeclaration, objectStateTypeDeclarationBlock);

                var objectStateFieldDeclaration = _encapsulateFieldCodeBuilder.BuildObjectStateFieldDeclaration(encapsulateFieldModel.ObjectStateUDTField);
                encapsulateFieldModel.NewContentAggregator.AddNewContent(NewContentType.DeclarationBlock, objectStateFieldDeclaration);
            }
        }

        private void ModifyReferences(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
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

        private void InsertNewContent(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
        {
            var encapsulateFieldInsertNewCodeModel = new EncapsulateFieldInsertNewCodeModel(model.SelectedFieldCandidates)
            {
                NewContentAggregator = model.NewContentAggregator,
            };

            _encapsulateFieldInsertNewCodeRefactoringAction.Refactor(encapsulateFieldInsertNewCodeModel, rewriteSession);
        }
    }
}
