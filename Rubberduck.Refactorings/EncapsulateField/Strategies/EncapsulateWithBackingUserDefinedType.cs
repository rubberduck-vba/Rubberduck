using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField.Strategies
{
    public interface IEncapsulateWithBackingUserDefinedType : IEncapsulateFieldStrategy
    {
        IEncapsulateFieldCandidate StateEncapsulationField { set; get; }
    }

    public class EncapsulateWithBackingUserDefinedType : EncapsulateFieldStrategiesBase, IEncapsulateWithBackingUserDefinedType
    {
        private const string DEFAULT_TYPE_IDENTIFIER = "This_Type";
        private const string DEFAULT_FIELD_IDENTIFIER = "this";


        public EncapsulateWithBackingUserDefinedType(QualifiedModuleName qmn, IIndenter indenter, IDeclarationFinderProvider declarationFinderProvider, IEncapsulateFieldNamesValidator validator)
            : base(qmn, indenter, declarationFinderProvider, validator)
        {
            StateEncapsulationField = _candidateFactory.CreateInsertableField(DEFAULT_FIELD_IDENTIFIER, DEFAULT_TYPE_IDENTIFIER, qmn, validator);
        }

        public IEncapsulateFieldCandidate StateEncapsulationField { set; get; }

        protected override void ModifyEncapsulatedVariable(IEncapsulateFieldCandidate target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            RewriterRemoveWorkAround.Remove(target.Declaration, rewriter);
            //rewriter.Remove(target.Declaration);
            return;
        }

        protected override EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields)
        {
            var nonUdtMemberFields = FlaggedEncapsulationFields
                    .Where(encFld => !encFld.IsUDTMember);

            var udt = new UDTDeclarationGenerator(StateEncapsulationField.AsTypeName);
            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                udt.AddMember(nonUdtMemberField);
            }
            newContent.AddDeclarationBlock(udt.TypeDeclarationBlock(Indenter));

            newContent.AddDeclarationBlock(udt.FieldDeclaration(StateEncapsulationField.NewFieldName));

            return newContent;
        }

        protected override EncapsulateFieldNewContent LoadNewPropertiesContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields, string postScript = null)
        {
            if (!FlaggedEncapsulationFields.Any()) { return newContent; }

            var udtMemberFields = FlaggedEncapsulationFields.Where(efd => efd.IsUDTMember);
            foreach (var udtMember in udtMemberFields)
            {
                newContent.AddCodeBlock(CreateUDTMemberProperty(udtMember));
            }

            return base.LoadNewPropertiesContent(newContent, FlaggedEncapsulationFields, postScript);
        }


        protected override IList<string> PropertiesContent(IEnumerable<IEncapsulateFieldCandidate> flaggedEncapsulationFields)
        {
            var textBlocks = new List<string>();
            foreach (var field in flaggedEncapsulationFields)
            {
                if (field is EncapsulatedUserDefinedTypeMember)
                {
                    continue;
                }
                textBlocks.Add(BuildPropertiesTextBlock(field.EncapsulationAttributes));
            }
            return textBlocks;
        }

        protected string BuildPropertiesTextBlock(IFieldEncapsulationAttributes attributes)
        {
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = $"{StateEncapsulationField.NewFieldName}.{attributes.PropertyName}",
                ParameterName = attributes.ParameterName,
                GenerateSetter = attributes.ImplementSetSetterType,
                GenerateLetter = attributes.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, Indenter.Indent(propertyTextLines, true));
        }

        private string CreateUDTMemberProperty(IEncapsulateFieldCandidate udtMember)
        {
            var parentField = UdtMemberTargetIDToParentMap[udtMember.TargetID];
            var generator = new PropertyGenerator
            {
                PropertyName = udtMember.PropertyName,
                AsTypeName = udtMember.AsTypeName,
                BackingField = $"{StateEncapsulationField.NewFieldName}.{parentField.PropertyName}.{udtMember.PropertyName}",
                ParameterName = udtMember.EncapsulationAttributes.ParameterName,
                GenerateSetter = udtMember.EncapsulationAttributes.ImplementSetSetterType,
                GenerateLetter = udtMember.EncapsulationAttributes.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, Indenter.Indent(propertyTextLines, true));
        }
    }
}
