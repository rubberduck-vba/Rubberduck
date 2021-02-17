using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using Rubberduck.Refactorings.ModifyUserDefinedType;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldUseBackingUDTMemberModel>
    {
        private readonly ICodeOnlyRefactoringAction<ModifyUserDefinedTypeModel> _modifyUDTRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;
        private readonly IEncapsulateFieldReferenceReplacer _referenceReplacer;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IReplacePrivateUDTMemberReferencesModelFactory replaceUDTMemberReferencesModelFactory,
            IEncapsulateFieldReferenceReplacerFactory encapsulateFieldReferenceReplacerFactory,
            IPropertyAttributeSetsGenerator propertyAttributeSetsGenerator,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory newContentAggregatorFactory)
                : base(rewritingManager)
        {
            _modifyUDTRefactoringAction = refactoringActionsProvider.ModifyUserDefinedType;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _newContentAggregatorFactory = newContentAggregatorFactory;

            _referenceReplacer = encapsulateFieldReferenceReplacerFactory.Create(replaceUDTMemberReferencesModelFactory,
                refactoringActionsProvider.ReplaceUDTMemberReferences,
                refactoringActionsProvider.ReplaceReferences,
                propertyAttributeSetsGenerator);
        }

        public override void Refactor(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any())
            {
                return;
            }

            ModifyFields(model, rewriteSession);

            _referenceReplacer.ReplaceReferences(model.SelectedFieldCandidates, rewriteSession, model.ObjectStateUDTField);

            InsertNewContent(model, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldUseBackingUDTMemberModel encapsulateFieldModel, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(encapsulateFieldModel.QualifiedModuleName);
            
            rewriter.RemoveVariables(encapsulateFieldModel.SelectedFieldCandidates.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());

            if (encapsulateFieldModel.ObjectStateUDTField.IsExistingDeclaration)
            {
                var model = new ModifyUserDefinedTypeModel(encapsulateFieldModel.ObjectStateUDTField.AsTypeDeclaration);

                foreach (var candidate in encapsulateFieldModel.SelectedFieldCandidates)
                {
                    model.AddNewMemberPrototype(candidate.Declaration, candidate.BackingIdentifier);
                }

                _modifyUDTRefactoringAction.Refactor(model, rewriteSession);
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
