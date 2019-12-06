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
            foreach (var field in model.FieldCandidates)
            {
                if (field is IEncapsulatedUserDefinedTypeField udt)
                {
                    udt.PropertyAccessExpression =
                        () =>
                        {
                            var accessor = udt.EncapsulateFlag ? udt.PropertyName : udt.NewFieldName;
                            return $"{StateUDTField.PropertyAccessExpression()}.{accessor}";
                        };

                    udt.ReferenceExpression = udt.PropertyAccessExpression;

                    foreach (var member in udt.Members)
                    {
                        member.PropertyAccessExpression = () => $"{udt.PropertyAccessExpression()}.{member.PropertyName}";
                        member.ReferenceExpression = () => $"{udt.PropertyAccessExpression()}.{member.PropertyName}";
                    }
                }
                else
                {
                    var efd = field;
                    efd.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{efd.PropertyName}";
                    efd.ReferenceExpression = efd.PropertyAccessExpression;
                }
            }

            foreach (var field in model.FlaggedFieldCandidates)
            {
                var attributes = field.EncapsulationAttributes;
                ModifyEncapsulatedField(field, /*attributes,*/ rewriteSession);
            }

            SetupReferenceModifications(model);
            foreach (var field in model.FlaggedFieldCandidates)
            {
                RenameReferences(field, rewriteSession);
                if (field is IEncapsulatedUserDefinedTypeField udtField)
                {
                    foreach (var member in udtField.Members)
                    {
                        RenameReferences(member, rewriteSession);
                    }
                }
            }

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);
            RewriterRemoveWorkAround.RemoveFieldsDeclaredInLists(rewriter);

            InsertNewContent(model.CodeSectionStartIndex, model, rewriteSession, asPreview);

            return rewriteSession;
        }

        public IEncapsulateFieldCandidate StateUDTField { set; get; }

        protected override void ModifyEncapsulatedField(IEncapsulateFieldCandidate target, /*IFieldEncapsulationAttributes attributes,*/ IRewriteSession rewriteSession)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            RewriterRemoveWorkAround.Remove(target.Declaration, rewriter);
            //rewriter.Remove(target.Declaration);
            return;
        }

        protected override EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> encapsulationCandidates)
        {
            var udt = new UDTDeclarationGenerator(StateUDTField.AsTypeName);

            var stateUDTMembers = encapsulationCandidates
                .Where(encFld => !encFld.IsUDTMember
                    && (encFld.EncapsulateFlag
                        || encFld is IEncapsulatedUserDefinedTypeField udtFld && udtFld.Members.Any(m => m.EncapsulateFlag)));

            udt.AddMembers(stateUDTMembers);

            newContent.AddDeclarationBlock(udt.TypeDeclarationBlock(Indenter));

            newContent.AddDeclarationBlock(udt.FieldDeclaration(StateUDTField.NewFieldName));

            return newContent;
        }
    }
}
