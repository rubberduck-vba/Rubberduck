using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveFieldsToUDT;
using Rubberduck.SmartIndenter;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : EncapsulateFieldRefactoringActionImplBase
    {
        private IObjectStateUDT _stateUDTField;
        private readonly DeclareFieldsAsUDTMembersRefactoringAction _convertFieldToUDTMemberRefactoringAction;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
            DeclareFieldsAsUDTMembersRefactoringAction convertFieldToUDTMemberRefactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            IIndenter indenter,
            IRewritingManager rewritingManager,
            ICodeBuilder codeBuilder)
                : base(declarationFinderProvider, indenter, rewritingManager, codeBuilder)
        {
            _convertFieldToUDTMemberRefactoringAction = convertFieldToUDTMemberRefactoringAction;
        }

        public override void Refactor(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            _stateUDTField = model.ObjectStateUDTField;

            RefactorImpl(model, rewriteSession);
        }

        protected override void ModifyFields(IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

            rewriter.RemoveVariables(SelectedFields.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());

            if (_stateUDTField.IsExistingDeclaration)
            {
                var model = new DeclareFieldsAsUDTMembersModel();

                foreach (var field in SelectedFields)
                {
                    model.AssignFieldToUserDefinedType(_stateUDTField.AsTypeDeclaration, field.Declaration as VariableDeclaration, field.PropertyIdentifier);
                }
                _convertFieldToUDTMemberRefactoringAction.Refactor(model, rewriteSession);
            }
        }

        protected override void LoadNewDeclarationBlocks()
        {
            if (_stateUDTField.IsExistingDeclaration)
            {
                return;
            }

            _stateUDTField.AddMembers(SelectedFields.Cast<IConvertToUDTMember>());

            AddContentBlock(NewContentType.TypeDeclarationBlock, _stateUDTField.TypeDeclarationBlock(_indenter));

            AddContentBlock(NewContentType.DeclarationBlock, _stateUDTField.FieldDeclarationBlock);
            return;
        }

        protected override void LoadFieldReferenceContextReplacements(IEncapsulateFieldCandidate field)
        {
            Debug.Assert(field is IConvertToUDTMember);

            var converted = field as IConvertToUDTMember;
            if (converted.WrappedCandidate is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
            {
                foreach (var member in udt.Members)
                {
                    foreach (var idRef in member.FieldContextReferences)
                    {
                        var replacementText = member.IdentifierForReference(idRef);
                        if (IsExternalReferenceRequiringModuleQualification(idRef))
                        {
                            replacementText = $"{udt.QualifiedModuleName.ComponentName}.{replacementText}";
                        }

                        SetUDTMemberReferenceRewriteContent(idRef, replacementText);
                    }
                }
            }
            else
            {
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementText = converted.IdentifierForReference(idRef);

                    if (IsExternalReferenceRequiringModuleQualification(idRef))
                    {
                        replacementText = $"{converted.QualifiedModuleName.ComponentName}.{replacementText}";
                    }

                    if (converted.Declaration.IsArray)
                    {
                        replacementText = $"{_stateUDTField.FieldIdentifier}.{replacementText}";
                    }

                    SetReferenceRewriteContent(idRef, replacementText);
                }
            }
        }
    }
}
