using Rubberduck.Parsing.Rewriter;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using Rubberduck.Refactorings.ModifyUserDefinedType;
using Rubberduck.Refactorings.DeleteDeclarations;

namespace Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldUseBackingUDTMemberModel>
    {
        private readonly ICodeOnlyRefactoringAction<ModifyUserDefinedTypeModel> _modifyUDTRefactoringAction;       
        private readonly ICodeOnlyRefactoringAction<EncapsulateFieldInsertNewCodeModel> _encapsulateFieldInsertNewCodeRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<DeleteDeclarationsModel> _deleteDeclarationsRefactoringAction;
        private readonly INewContentAggregatorFactory _newContentAggregatorFactory;
        private readonly IEncapsulateFieldReferenceReplacerFactory _encapsulateFieldReferenceReplacerFactory;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
            IEncapsulateFieldRefactoringActionsProvider refactoringActionsProvider,
            IEncapsulateFieldReferenceReplacerFactory encapsulateFieldReferenceReplacerFactory,
            IRewritingManager rewritingManager,
            INewContentAggregatorFactory newContentAggregatorFactory)
                : base(rewritingManager)
        {
            _modifyUDTRefactoringAction = refactoringActionsProvider.ModifyUserDefinedType;
            _encapsulateFieldInsertNewCodeRefactoringAction = refactoringActionsProvider.EncapsulateFieldInsertNewCode;
            _deleteDeclarationsRefactoringAction = refactoringActionsProvider.DeleteDeclarations;
            _encapsulateFieldReferenceReplacerFactory = encapsulateFieldReferenceReplacerFactory;
            _newContentAggregatorFactory = newContentAggregatorFactory;
        }

        public override void Refactor(EncapsulateFieldUseBackingUDTMemberModel model, IRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any())
            {
                return;
            }

            ModifyFields(model, rewriteSession);

            var referenceReplacer = _encapsulateFieldReferenceReplacerFactory.Create();
            referenceReplacer.ReplaceReferences(model.SelectedFieldCandidates, rewriteSession);

            InsertNewContent(model, rewriteSession);
        }

        private void ModifyFields(EncapsulateFieldUseBackingUDTMemberModel encapsulateFieldModel, IRewriteSession rewriteSession)
        {
            var deletionsModel = new DeleteDeclarationsModel(encapsulateFieldModel.SelectedFieldCandidates.Select(f => f.Declaration));
            
            _deleteDeclarationsRefactoringAction.Refactor(deletionsModel, rewriteSession);

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
