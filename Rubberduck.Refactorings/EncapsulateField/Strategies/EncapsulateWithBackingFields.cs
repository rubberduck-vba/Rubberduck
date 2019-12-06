using Rubberduck.Parsing.Grammar;
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
    public class EncapsulateWithBackingFields : EncapsulateFieldStrategiesBase
    {
        public EncapsulateWithBackingFields(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldNamesValidator validator)
            : base(qmn, indenter, validator)
        {
        }

        protected override void ModifyEncapsulatedField(IEncapsulateFieldCandidate field, IRewriteSession rewriteSession)
        {
            var attributes = field.EncapsulationAttributes;
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            if (field.Declaration.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(field.Declaration.IdentifierName))
            {
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                return;
            }

            if (field.Declaration.IsDeclaredInList())
            {
                RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
                //rewriter.Remove(target.Declaration);
            }
            else
            {
                rewriter.Rename(field.Declaration, attributes.NewFieldName);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
            return;
        }

        protected override EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> encapsulationCandates)
        {
            var nonUdtMemberFields = encapsulationCandates
                    .Where(encFld => !encFld.IsUDTMember && encFld.EncapsulateFlag);

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;

                if (nonUdtMemberField.Declaration.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(nonUdtMemberField.Declaration.IdentifierName))
                {
                    continue;
                }

                if (nonUdtMemberField.Declaration.IsDeclaredInList())
                {
                    var targetIdentifier = nonUdtMemberField.Declaration.Context.GetText().Replace(attributes.IdentifierName, attributes.NewFieldName);
                    var newField = nonUdtMemberField.Declaration.IsTypeSpecified ? $"{Tokens.Private} {targetIdentifier}" : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {nonUdtMemberField.Declaration.AsTypeName}";

                    newContent.AddDeclarationBlock(newField);
                }
            }
            return newContent;
        }
    }
}
