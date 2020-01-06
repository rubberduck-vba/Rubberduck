using Antlr4.Runtime;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
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
            model.AssignCandidateValidations(EncapsulateFieldStrategy.ConvertFieldsToUDTMembers);
            _convertedFields = new List<IConvertToUDTMember>();
            if (File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt"))
            {
                foreach (var field in model.SelectedFieldCandidates)
                {
                    _convertedFields.Add(new ConvertToUDTMember(field, model.StateUDTField));
                }
            }
            else
            {
                _convertedFields = model.SelectedFieldCandidates.Cast<IConvertToUDTMember>().ToList();
            }
            foreach (var convert in _convertedFields)
            {
                convert.NameValidator = convert.Declaration.IsArray
                    ? model.ValidatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMemberArray)
                    : model.ValidatorProvider.NameOnlyValidator(Validators.UserDefinedTypeMember);

                convert.ConflictFinder = model.ValidatorProvider.ConflictDetector(EncapsulateFieldStrategy.ConvertFieldsToUDTMembers, declarationFinderProvider);
            }
            _stateUDTField = model.StateUDTField;
        }

        protected override void ModifyFields(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            foreach (var field in  model.SelectedFieldCandidates)
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
            foreach (var field in model.SelectedFieldCandidates)
            {
                field.LoadFieldReferenceContextReplacements(_stateUDTField.FieldIdentifier);
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
            var propertyGenerationSpecs = _convertedFields // model.SelectedFieldCandidates
                                                .SelectMany(f => f.PropertyAttributeSets);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddContentBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, _indenter));
            }
        }
    }
}
