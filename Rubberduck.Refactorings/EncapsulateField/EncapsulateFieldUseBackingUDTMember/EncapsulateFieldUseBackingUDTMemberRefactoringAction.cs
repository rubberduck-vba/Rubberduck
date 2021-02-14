using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using System.Collections.Generic;
using Rubberduck.Refactorings.ModifyUserDefinedType;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldUseBackingUDTMemberModel>
    {
        private readonly ICodeOnlyRefactoringAction<ModifyUserDefinedTypeModel> _modifyUDTRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplacePrivateUDTMemberReferencesModel> _replacePrivateUDTMemberReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<ReplaceReferencesModel> _replaceReferencesRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;
        private readonly IReplacePrivateUDTMemberReferencesModelFactory _replaceUDTMemberReferencesModelFactory;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IReplacePrivateUDTMemberReferencesModelFactory replaceUDTMemberReferencesModelFactory,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory newContentAggregatorFactory)
                : base(rewritingManager)
        {
            _modifyUDTRefactoringAction = refactoringActionsProvider.ModifyUserDefinedType;
            _replacePrivateUDTMemberReferencesRefactoringAction = refactoringActionsProvider.ReplaceUDTMemberReferences;
            _replaceReferencesRefactoringAction = refactoringActionsProvider.ReplaceReferences;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _replaceUDTMemberReferencesModelFactory = replaceUDTMemberReferencesModelFactory;
            _newContentAggregatorFactory = newContentAggregatorFactory;
        }

        public override void Refactor(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any())
            {
                return;
            }

            ModifyFields(model, rewriteSession);

            ModifyReferences(model, rewriteSession);

            InsertNewContent(model, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldUseBackingUDTMemberModel encapsulateFieldModel, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(encapsulateFieldModel.QualifiedModuleName);

            if (encapsulateFieldModel.ObjectStateUDTField.IsExistingDeclaration)
            {
                var model = new ModifyUserDefinedTypeModel(encapsulateFieldModel.ObjectStateUDTField.AsTypeDeclaration);

                foreach (var candidate in encapsulateFieldModel.SelectedFieldCandidates)
                {
                    model.AddNewMemberPrototype(candidate.Declaration, candidate.BackingIdentifier);
                }

                _modifyUDTRefactoringAction.Refactor(model,rewriteSession);
            }

            rewriter.RemoveVariables(encapsulateFieldModel.SelectedFieldCandidates.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());
        }

        private void ModifyReferences(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
        {
            var privateUDTFields = model.SelectedFieldCandidates.Cast<IEncapsulateFieldAsUDTMemberCandidate>()
                .Where(f => f.WrappedCandidate is IUserDefinedTypeCandidate udt
                    && udt.Declaration.AsTypeDeclaration.Accessibility == Accessibility.Private);

            ReplaceMemberReferencesOfPrivateUDTFields(model, privateUDTFields.Select(p => (p, p.WrappedCandidate as IUserDefinedTypeCandidate)), rewriteSession);

            ReplaceEncapsulatedFieldReferences(model.SelectedFieldCandidates.Except(privateUDTFields), model.ObjectStateUDTField, rewriteSession);
        }

        private void ReplaceMemberReferencesOfPrivateUDTFields(EncapsulateFieldUseBackingUDTMemberModel model, IEnumerable<(IEncapsulateFieldAsUDTMemberCandidate candidate, IUserDefinedTypeCandidate wrappedUDT)> privateUDTTuples, IRewriteSession rewriteSession)
        {
            if (!privateUDTTuples.Any())
            {
                return;
            }

            var replacePrivateUDTMemberReferencesModel 
                = _replaceUDTMemberReferencesModelFactory.Create(privateUDTTuples.Select(f => f.candidate.Declaration).Cast<VariableDeclaration>());

            foreach (var udtTuple in privateUDTTuples)
            {
               InitializeModel(model, replacePrivateUDTMemberReferencesModel, udtTuple);
            }

            _replacePrivateUDTMemberReferencesRefactoringAction.Refactor(replacePrivateUDTMemberReferencesModel, rewriteSession);
        }

        private void ReplaceEncapsulatedFieldReferences(IEnumerable<IEncapsulateFieldCandidate> nonPrivateUDTFields, IObjectStateUDT objectStateUDTField, IRewriteSession rewriteSession)
        {
            if (!nonPrivateUDTFields.Any())
            {
                return;
            }

            var replaceReferencesModel = new ReplaceReferencesModel()
            {
                ModuleQualifyExternalReferences = true,
            };

            foreach (var field in nonPrivateUDTFields)
            {
                InitializeModel(replaceReferencesModel, field, objectStateUDTField);
            }

            _replaceReferencesRefactoringAction.Refactor(replaceReferencesModel, rewriteSession);
        }

        private void InitializeModel(EncapsulateFieldUseBackingUDTMemberModel model, ReplacePrivateUDTMemberReferencesModel replacePrivateUDTMemberReferencesModel, (IEncapsulateFieldAsUDTMemberCandidate candidate, IUserDefinedTypeCandidate wrappedUDT) privateUDTTuple)
        {
            var memberToLocalExpression = new Dictionary<Declaration, string>();
            foreach (var member in privateUDTTuple.wrappedUDT.Members)
            {
                var memberAccessExpression = privateUDTTuple.candidate.IsReadOnly 
                    ? model.GetLocalBackingExpression(privateUDTTuple.candidate, member) 
                    : member.PropertyIdentifier;

                memberToLocalExpression.Add(member.Declaration, memberAccessExpression);
            }

            foreach (var udtMember in replacePrivateUDTMemberReferencesModel.UDTMembers)
            {
                var udtExpressions = new PrivateUDTMemberReferenceReplacementExpressions($"{privateUDTTuple.candidate.IdentifierName}.{udtMember.IdentifierName}")
                {
                    UDTMemberInternalReferenceExpression = memberToLocalExpression[udtMember]
                };

                replacePrivateUDTMemberReferencesModel.AssignUDTMemberReferenceExpressions(privateUDTTuple.candidate.Declaration as VariableDeclaration, udtMember, udtExpressions);
            }
        }

        private void InitializeModel(ReplaceReferencesModel model, IEncapsulateFieldCandidate field, IObjectStateUDT objectStateUDTField)
        {
            foreach (var idRef in field.Declaration.References)
            {
                var replacementExpression = field.PropertyIdentifier;

                if (idRef.QualifiedModuleName == field.QualifiedModuleName && (field.Declaration.IsArray || field.IsReadOnly))
                {
                    replacementExpression = $"{objectStateUDTField.FieldIdentifier}.{field.BackingIdentifier}";
                }

                model.AssignReferenceReplacementExpression(idRef, replacementExpression);
            }
        }

        private void InsertNewContent(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
        {
            var aggregator = model.NewContentAggregator ?? _newContentAggregatorFactory.Create();
            model.NewContentAggregator = null;

            var encapsulateFieldInsertNewCodeModel = new EncapsulateFieldInsertNewCodeModel(model.SelectedFieldCandidates)
            {
                NewContentAggregator = aggregator,
                ObjectStateUDTField = model.ObjectStateUDTField
            };

            _encapsulateFieldInsertNewCodeRefactoringAction.Refactor(encapsulateFieldInsertNewCodeModel, rewriteSession);
        }
    }
}
