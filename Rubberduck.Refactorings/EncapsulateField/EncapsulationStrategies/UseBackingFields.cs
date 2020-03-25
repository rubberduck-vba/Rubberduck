using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UseBackingFields : EncapsulateFieldStrategyBase
    {
        public UseBackingFields(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter)
            : base(declarationFinderProvider, model, indenter){ }

        protected override void ModifyFields(IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            foreach (var field in SelectedFields)
            {
                if (field.Declaration.HasPrivateAccessibility() && field.BackingIdentifier.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                if (field.Declaration.IsDeclaredInList() && !field.Declaration.HasPrivateAccessibility())
                {
                    refactorRewriteSession.Remove(field.Declaration, rewriter);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.BackingIdentifier);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
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
            //New field declarations created here were removed from their list within ModifyFields(...)
            var fieldsRequiringNewDeclaration = SelectedFields
                .Where(field => field.Declaration.IsDeclaredInList()
                                    && field.Declaration.Accessibility != Accessibility.Private);

            foreach (var field in fieldsRequiringNewDeclaration)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.BackingIdentifier);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddContentBlock(NewContentTypes.DeclarationBlock, newField);
            }
        }
    }
}
