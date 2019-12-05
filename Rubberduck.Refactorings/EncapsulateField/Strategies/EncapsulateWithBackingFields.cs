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

        protected override void ModifyEncapsulatedVariable(IEncapsulateFieldCandidate target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession)
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

        protected override IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            foreach (var field in model.UDTFieldCandidates)
            {
                if (!field.TypeDeclarationIsPrivate)
                {
                    field.ReferenceExpression = () => field.PropertyName;
                }
                else
                {
                    foreach (var member in field.Members)
                    {
                        member.PropertyAccessExpression = () => $"{field.PropertyAccessExpression()}.{member.IdentifierName}";
                    }
                }
            }

            foreach (var field in model.FieldCandidates.Except(model.UDTFieldCandidates).Where(fld => fld.EncapsulateFlag))
            {
                field.ReferenceExpression = () => field.PropertyName;
            }

            return base.RefactorRewrite(model, rewriteSession, asPreview);
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
                    var targetIdentifier = nonUdtMemberField.Declaration.Context.GetText().Replace(attributes.Identifier, attributes.NewFieldName);
                    var newField = nonUdtMemberField.Declaration.IsTypeSpecified ? $"{Tokens.Private} {targetIdentifier}" : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {nonUdtMemberField.Declaration.AsTypeName}";

                    newContent.AddDeclarationBlock(newField);
                }
            }
            return newContent;
        }
    }
}
