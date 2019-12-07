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
        IEncapsulateFieldCandidate StateUDTField { set; get; }
    }

    public class EncapsulateWithBackingUserDefinedType : EncapsulateFieldStrategiesBase, IEncapsulateWithBackingUserDefinedType
    {
        public EncapsulateWithBackingUserDefinedType(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldNamesValidator validator)
            : base(qmn, indenter, validator) { }

        protected override IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            foreach (var field in model.FlaggedFieldCandidates)
            {
                if (field is IEncapsulatedUserDefinedTypeField udt)
                {
                    udt.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{udt.PropertyName}";

                    udt.ReferenceExpression = udt.PropertyAccessExpression;

                    foreach (var member in udt.Members)
                    {
                        member.PropertyAccessExpression = () => $"{udt.PropertyAccessExpression()}.{member.PropertyName}";
                        member.ReferenceExpression = () => $"{udt.PropertyAccessExpression()}.{member.PropertyName}";
                    }
                }
                else
                {
                    field.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{field.PropertyName}";
                    field.ReferenceExpression = field.PropertyAccessExpression;
                }
            }

            return base.RefactorRewrite(model, rewriteSession, asPreview);
        }

        public IEncapsulateFieldCandidate StateUDTField { set; get; }

        protected override void ModifyField(IEncapsulateFieldCandidate field, IRewriteSession rewriteSession)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
            //rewriter.Remove(target.Declaration);
            return;
        }

        protected override EncapsulateFieldNewContent LoadNewDeclarationBlocks(EncapsulateFieldNewContent newContent, EncapsulateFieldModel model)
        {
            var udt = new UDTDeclarationGenerator(StateUDTField.AsTypeName);

            udt.AddMembers(model.FlaggedFieldCandidates);

            newContent.AddTypeDeclarationBlock(udt.TypeDeclarationBlock(Indenter));

            newContent.AddDeclarationBlock(udt.FieldDeclaration(StateUDTField.NewFieldName));

            return newContent;
        }
    }
}
