using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.SmartIndenter;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class UseBackingFields : EncapsulateFieldStrategyBase
    {
        private IEnumerable<IEncapsulateFieldCandidate> _fieldsToDeleteAndReplace;

        public UseBackingFields(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter, ICodeBuilder codeBuilder)
            : base(declarationFinderProvider, model, indenter, codeBuilder)
        {
            _fieldsToDeleteAndReplace = SelectedFields.Where(f => f.Declaration.IsDeclaredInList() && !f.Declaration.HasPrivateAccessibility()).ToList();
        }


        protected override void ModifyFields(IRewriteSession refactorRewriteSession)
        {
            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);

            rewriter.RemoveVariables(_fieldsToDeleteAndReplace.Select(f => f.Declaration).Cast<VariableDeclaration>());

            foreach (var field in SelectedFields.Except(_fieldsToDeleteAndReplace))
            {
                if (field.Declaration.HasPrivateAccessibility() && field.BackingIdentifier.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.BackingIdentifier);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
        }

        protected override void ModifyReferences(IRewriteSession refactorRewriteSession)
        {
            foreach (var field in SelectedFields)
            {
                LoadFieldReferenceContextReplacements(field);
            }

            RewriteReferences(refactorRewriteSession);
        }

        protected override void LoadNewDeclarationBlocks()
        {
            //New field declarations created here were removed from their 
            //variable list statement within ModifyFields(...)
            foreach (var field in _fieldsToDeleteAndReplace)
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
