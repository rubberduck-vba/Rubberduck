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
        public EncapsulateWithBackingFields(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldValidator validator)
            : base(qmn, indenter, validator) { }

        protected override void ModifyFields(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            foreach (var field in model.SelectedFieldCandidates)
            {
                var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

                if (field.Declaration.Accessibility == Accessibility.Private && field.NewFieldName.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                if (field.Declaration.IsDeclaredInList())
                {
                    RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
                    //rewriter.Remove(target.Declaration);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.NewFieldName);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
        }

        protected override void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
        {
            //New field declarations created here were removed from their list within ModifyFields(...)
            var fieldsRequiringNewDeclaration = model.SelectedFieldCandidates
                .Where(field => field.Declaration.IsDeclaredInList()
                                    && field.Declaration.Accessibility != Accessibility.Private);

            foreach (var field in fieldsRequiringNewDeclaration)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.NewFieldName);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddCodeBlock(NewContentTypes.DeclarationBlock, newField);
            }
        }
    }
}
