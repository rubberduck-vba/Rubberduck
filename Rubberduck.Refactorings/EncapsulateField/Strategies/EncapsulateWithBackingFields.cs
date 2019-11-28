using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField.Strategies
{
    public class EncapsulateWithBackingFields : EncapsulateFieldStrategiesBase
    {
        public EncapsulateWithBackingFields(QualifiedModuleName qmn, IIndenter indenter)
            : base(qmn, indenter)
        {

        }

        protected override void ModifyEncapsulatedVariable(IEncapsulateFieldCandidate target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession) //, bool asUDT = false) //, EncapsulateFieldNewContent newContent)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            if (target.Declaration.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(target.Declaration.IdentifierName))
            {
                rewriter.MakeImplicitDeclarationTypeExplicit(target.Declaration);
                return;
            }

            if (target.Declaration.IsDeclaredInList())
            {
                RewriterRemoveWorkAround.Remove(target.Declaration, rewriter);
                //rewriter.Remove(target.Declaration);
            }
            else
            {
                rewriter.Rename(target.Declaration, attributes.NewFieldName);
                rewriter.SetVariableVisiblity(target.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(target.Declaration);
            }
            return;
        }

        protected override EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields)
        {
            var nonUdtMemberFields = FlaggedEncapsulationFields
                    .Where(encFld => encFld.Declaration.IsVariable());

            //if (EncapsulateWithUDT)
            //{
            //    var udt = new UDTDeclarationGenerator(EncapsulateWithUDT_TypeIdentifier, _indenter);
            //    foreach (var nonUdtMemberField in nonUdtMemberFields)
            //    {
            //        udt.AddMember(nonUdtMemberField);
            //    }
            //    newContent.AddDeclarationBlock(udt.TypeDeclarationBlock);
            //    newContent.AddDeclarationBlock(udt.FieldDeclaration(EncapsulateWithUDT_FieldName));

            //    var udtMemberFields = FlaggedEncapsulationFields.Where(efd => efd.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember));
            //    foreach (var udtMember in udtMemberFields)
            //    {
            //        newContent.AddCodeBlock(EncapsulateInUDT_UDTMemberProperty(udtMember));
            //    }

            //    return newContent;
            //}

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;

                if (nonUdtMemberField.Declaration.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(nonUdtMemberField.Declaration.IdentifierName))
                {
                    continue;
                }

                if (nonUdtMemberField.Declaration.IsDeclaredInList())
                {
                    var targetIdentifier = nonUdtMemberField.Declaration.Context.GetText().Replace(attributes.TargetName, attributes.NewFieldName);
                    var newField = nonUdtMemberField.Declaration.IsTypeSpecified ? $"{Tokens.Private} {targetIdentifier}" : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {nonUdtMemberField.Declaration.AsTypeName}";

                    newContent.AddDeclarationBlock(newField);
                }
            }
            return newContent;
        }

        protected override IList<string> PropertiesContent(IEnumerable<IEncapsulateFieldCandidate> flaggedEncapsulationFields)
        {
            //get
            //{
                var textBlocks = new List<string>();
                foreach (var field in flaggedEncapsulationFields)
                {
                    //if (/*EncapsulateWithUDT &&*/ field is EncapsulatedUserDefinedTypeMember)
                    //{
                    //    continue;
                    //}
                    textBlocks.Add(BuildPropertiesTextBlock(field.EncapsulationAttributes));
                }
                return textBlocks;
            //}
        }

        private string BuildPropertiesTextBlock(IFieldEncapsulationAttributes attributes)
        {
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                //BackingField = EncapsulateWithUDT
                //                    ? $"{EncapsulateWithUDT_FieldName}.{attributes.PropertyName}"
                //                    : attributes.FieldReferenceExpression,
                BackingField = attributes.FieldReferenceExpression,
                ParameterName = attributes.ParameterName,
                GenerateSetter = attributes.ImplementSetSetterType,
                GenerateLetter = attributes.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, Indenter.Indent(propertyTextLines, true));
        }
    }
}
