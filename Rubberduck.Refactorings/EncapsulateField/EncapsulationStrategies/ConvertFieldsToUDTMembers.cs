using Antlr4.Runtime;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class ConvertFieldsToUDTMembers : EncapsulateFieldStrategyBase
    {
        private List<IConvertToUDTMember> _convertedFields;
        private IObjectStateUDT _stateUDTField;

        public ConvertFieldsToUDTMembers(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter)
            : base(declarationFinderProvider, model, indenter)
        {
            _convertedFields = new List<IConvertToUDTMember>();
            foreach (var field in model.SelectedFieldCandidates)
            {
                _convertedFields.Add(new ConvertToUDTMember(field, model.StateUDTField));
            }
            _stateUDTField = model.StateUDTField;
        }

        protected override void ModifyFields(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            foreach (var field in _convertedFields)
            {
                refactorRewriteSession.Remove(field.Declaration, rewriter);
            }

            if (_stateUDTField.IsExistingDeclaration)
            {
                _stateUDTField.AddMembers(_convertedFields);

                rewriter.Replace(_stateUDTField.AsTypeDeclaration, _stateUDTField.TypeDeclarationBlock(_indenter));
            }
        }

        protected override void ModifyReferences(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            foreach (var field in _convertedFields)
            {
                LoadFieldReferenceContextReplacements(field);
            }

            RewriteReferences(model, refactorRewriteSession);
        }

        protected override void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
        {
            if (_stateUDTField.IsExistingDeclaration) { return; }

            _stateUDTField.AddMembers(_convertedFields);

            AddContentBlock(NewContentTypes.TypeDeclarationBlock, _stateUDTField.TypeDeclarationBlock(_indenter));

            AddContentBlock(NewContentTypes.DeclarationBlock, _stateUDTField.FieldDeclarationBlock);
            return;
        }

        protected override void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = _convertedFields.SelectMany(f => f.PropertyAttributeSets);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddContentBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, _indenter));
            }
        }

        protected override void LoadFieldReferenceContextReplacements(IEncapsulatableField field)
        {
            Debug.Assert(field is IConvertToUDTMember);

            var converted = field as IConvertToUDTMember;
            if (converted.WrappedCandidate is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
            {
                foreach (var member in udt.Members)
                {
                    foreach (var idRef in member.ParentContextReferences)
                    {
                        var replacementText = member.ReferenceAccessor(idRef);
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
                    var replacementText = converted.ReferenceAccessor(idRef);

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
