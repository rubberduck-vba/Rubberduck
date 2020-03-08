using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
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
        private IObjectStateUDT _stateUDTField;

        public ConvertFieldsToUDTMembers(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter)
            : base(declarationFinderProvider, model, indenter)
        {
            _stateUDTField = model.ObjectStateUDTField;
        }

        protected override void ModifyFields(IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            foreach (var field in SelectedFields)
            {
                refactorRewriteSession.Remove(field.Declaration, rewriter);
            }

            if (_stateUDTField.IsExistingDeclaration)
            {
                _stateUDTField.AddMembers(SelectedFields.Cast<IConvertToUDTMember>());

                rewriter.Replace(_stateUDTField.AsTypeDeclaration, _stateUDTField.TypeDeclarationBlock(_indenter));
            }
        }

        protected override void ModifyReferences(IEncapsulateFieldRewriteSession refactorRewriteSession)
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

        protected override void LoadNewPropertyBlocks()
        {
            var propertyGenerationSpecs = SelectedFields.SelectMany(f => f.PropertyAttributeSets);

            foreach (var selectedField in SelectedFields)
            {
                var converted = selectedField as IConvertToUDTMember;
                foreach (var set in selectedField.PropertyAttributeSets)
                {
                    if (converted.Declaration is VariableDeclaration variableDeclaration)
                    {
                        var getContent = $"{set.PropertyName} = {set.BackingField}";
                        if (set.UsesSetAssignment)
                        {
                            getContent = $"{Tokens.Set} {getContent}";
                        }
                        AddContentBlock(NewContentTypes.MethodBlock, variableDeclaration.FieldToPropertyBlock(DeclarationType.PropertyGet, set.PropertyName, content: $"{_defaultIndent}{getContent}"));
                        if (converted.IsReadOnly)
                        {
                            continue;
                        }
                        if (set.GenerateLetter)
                        {
                            AddContentBlock(NewContentTypes.MethodBlock, variableDeclaration.FieldToPropertyBlock(DeclarationType.PropertyLet, set.PropertyName, content: $"{_defaultIndent}{set.BackingField} = {set.ParameterName}"));
                        }
                        if (set.GenerateSetter)
                        {
                            AddContentBlock(NewContentTypes.MethodBlock, variableDeclaration.FieldToPropertyBlock(DeclarationType.PropertySet, set.PropertyName, content: $"{_defaultIndent}{Tokens.Set} {set.BackingField} = {set.ParameterName}"));
                        }
                    }
                }
            }
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
