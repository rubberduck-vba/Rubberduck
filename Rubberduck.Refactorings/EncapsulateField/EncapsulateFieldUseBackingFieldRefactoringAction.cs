using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.SmartIndenter;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldUseBackingFieldRefactoringAction : EncapsulateFieldRefactoringActionImplBase
    {
        public EncapsulateFieldUseBackingFieldRefactoringAction(
                IDeclarationFinderProvider declarationFinderProvider,
                IIndenter indenter,
                IRewritingManager rewritingManager,
                ICodeBuilder codeBuilder)
            : base(declarationFinderProvider, indenter, rewritingManager, codeBuilder)
        {}

        public override void Refactor(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            RefactorImpl(model, rewriteSession);
        }

        protected override void ModifyFields(IRewriteSession rewriteSession)
        {
            var fieldsToDeleteAndReplace = SelectedFields.Where(f => IsFieldToDeleteAndReplace(f));
            var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

            RemoveFields(fieldsToDeleteAndReplace.Select(f => f.Declaration), rewriteSession);

            foreach (var field in SelectedFields.Except(fieldsToDeleteAndReplace))
            {
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);

                if (!field.Declaration.HasPrivateAccessibility())
                {
                    rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                }

                if (!field.BackingIdentifier.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.Rename(field.Declaration, field.BackingIdentifier);
                }
            }
        }

        protected override void ModifyReferences(IRewriteSession rewriteSession)
        {
            foreach (var field in SelectedFields)
            {
                LoadFieldReferenceContextReplacements(field);
            }

            RewriteReferences(rewriteSession);
        }

        protected override void LoadNewDeclarationBlocks()
        {
            //Fields to create here were deleted in ModifyFields(...)
            foreach (var field in SelectedFields.Where(f => IsFieldToDeleteAndReplace(f)))
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.BackingIdentifier);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddContentBlock(NewContentType.DeclarationBlock, newField);
            }
        }

        private static bool IsFieldToDeleteAndReplace(IEncapsulateFieldCandidate field)
            => field.Declaration.IsDeclaredInList() && !field.Declaration.HasPrivateAccessibility();
    }
}
