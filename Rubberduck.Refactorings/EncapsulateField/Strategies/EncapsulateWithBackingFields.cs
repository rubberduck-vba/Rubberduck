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

        protected override void ModifyField(IEncapsulateFieldCandidate field, IRewriteSession rewriteSession)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            if (field.Declaration.Accessibility == Accessibility.Private && field.NewFieldName.Equals(field.Declaration.IdentifierName))
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
                rewriter.Rename(field.Declaration, field.NewFieldName);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
            return;
        }

        protected override EncapsulateFieldNewContent LoadNewDeclarationBlocks(EncapsulateFieldNewContent newContent, EncapsulateFieldModel model)
        {
            foreach (var field in model.FlaggedFieldCandidates)
            {

                if (field.Declaration.Accessibility == Accessibility.Private && field.NewFieldName.Equals(field.Declaration.IdentifierName))
                {
                    continue;
                }

                //Fields within a list (where Accessibility is 'Public' 
                //are removed from the list (within ModifyField(...)) and 
                //inserted as a new Declaration
                if (field.Declaration.IsDeclaredInList())
                {
                    var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.NewFieldName);
                    var newField = field.Declaration.IsTypeSpecified 
                        ? $"{Tokens.Private} {targetIdentifier}" 
                        : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                    newContent.AddDeclarationBlock(newField);
                }
            }
            return newContent;
        }
    }
}
