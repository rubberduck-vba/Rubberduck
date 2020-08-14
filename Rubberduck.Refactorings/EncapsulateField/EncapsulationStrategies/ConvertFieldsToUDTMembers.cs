using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.SmartIndenter;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class ConvertFieldsToUDTMembers : EncapsulateFieldStrategyBase
    {
        private IObjectStateUDT _stateUDTField;

        public ConvertFieldsToUDTMembers(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter, ICodeBuilder codeBuilder)
            : base(declarationFinderProvider, model, indenter, codeBuilder)
        {
            _stateUDTField = model.ObjectStateUDTField;
        }

        protected override void ModifyFields(IRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            rewriter.RemoveVariables(SelectedFields.Select(f => f.Declaration)
                .Cast<VariableDeclaration>());

            if (_stateUDTField.IsExistingDeclaration)
            {
                _stateUDTField.AddMembers(SelectedFields.Cast<IConvertToUDTMember>());

                rewriter.Replace(_stateUDTField.AsTypeDeclaration, _stateUDTField.TypeDeclarationBlock(_indenter));
            }
        }

        protected override void ModifyReferences(IRewriteSession refactorRewriteSession)
        {
            foreach (var field in SelectedFields)
            {
                LoadFieldReferenceContextReplacements(field);
            }

            RewriteReferences(refactorRewriteSession);
        }

        protected override void LoadNewDeclarationBlocks()
        {
            if (_stateUDTField.IsExistingDeclaration) { return; }

            _stateUDTField.AddMembers(SelectedFields.Cast<IConvertToUDTMember>());

            AddContentBlock(NewContentTypes.TypeDeclarationBlock, _stateUDTField.TypeDeclarationBlock(_indenter));

            AddContentBlock(NewContentTypes.DeclarationBlock, _stateUDTField.FieldDeclarationBlock);
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
