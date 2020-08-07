using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldUseBackingUDTMemberRefactoringAction : EncapsulateFieldRefactoringActionImplBase
    {
        private IObjectStateUDT _stateUDTField;

        public EncapsulateFieldUseBackingUDTMemberRefactoringAction(
                IDeclarationFinderProvider declarationFinderProvider,
                IIndenter indenter,
                IRewritingManager rewritingManager,
                ICodeBuilder codeBuilder)
            : base(declarationFinderProvider, indenter, rewritingManager, codeBuilder)
        {}

        public override void Refactor(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            _stateUDTField = model.ObjectStateUDTField;

            RefactorImpl(model, rewriteSession);
        }

        protected override void ModifyFields(IRewriteSession rewriteSession)
        {
            RemoveFields(SelectedFields.Select(sf => sf.Declaration), rewriteSession);

            if (_stateUDTField.IsExistingDeclaration)
            {
                _stateUDTField.AddMembers(SelectedFields.Cast<IConvertToUDTMember>());

                var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);
                rewriter.Replace(_stateUDTField.AsTypeDeclaration, _stateUDTField.TypeDeclarationBlock(_indenter));
            }
        }

        protected override void ModifyReferences(IRewriteSession rewriteSession)
        {
            foreach (var field in SelectedFields)
            {
                LoadFieldReferenceContextReplacements(field);
            }

            RewriteReferences(rewriteSession);
        }

        protected override void LoadNewDeclarationBlocks()
        {
            if (_stateUDTField.IsExistingDeclaration) { return; }

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
