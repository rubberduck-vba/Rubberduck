using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.SmartIndenter;
using System;
using System.Collections.Generic;
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
            var fieldDeclarationsToDeleteAndReplace = SelectedFields.Where(f => IsFieldToDeleteAndReplace(f));
            var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

            rewriter.RemoveVariables(fieldDeclarationsToDeleteAndReplace.Select(f => f.Declaration).Cast<VariableDeclaration>());

            var fieldDeclaraionsToRetain = SelectedFields.Except(fieldDeclarationsToDeleteAndReplace).ToList();

            if (fieldDeclaraionsToRetain.Any())
            {
                MakeImplicitDeclarationTypeExplicit(fieldDeclaraionsToRetain, rewriter);


                SetPrivateVariableVisiblity(fieldDeclaraionsToRetain, rewriter);

                Rename(fieldDeclaraionsToRetain, rewriter);
            }
        }

        private static void MakeImplicitDeclarationTypeExplicit(IEnumerable<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
        {
            foreach (var element in fields.Select(f => f.Declaration))
            {
                if (!element.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
                {
                    rewriter.InsertAfter(element.Context.Stop.TokenIndex, $" {Tokens.As} {element.AsTypeName}");
                }
            }
        }

        private static void SetPrivateVariableVisiblity(IEnumerable<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
        {
            var visibility = Accessibility.Private.TokenString();
            foreach (var element in fields.Where(f => !f.Declaration.HasPrivateAccessibility()).Select(f => f.Declaration))
            {
                if (!element.IsVariable())
                {
                    throw new ArgumentException();
                }

                var variableStmtContext = element.Context.GetAncestor<VBAParser.VariableStmtContext>();
                var visibilityContext = variableStmtContext.GetChild<VBAParser.VisibilityContext>();

                if (visibilityContext != null)
                {
                    rewriter.Replace(visibilityContext, visibility);
                    continue;
                }
                rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{visibility} ");
            }
        }

        private static void Rename(IEnumerable<IEncapsulateFieldCandidate> fields, IModuleRewriter rewriter)
        {
            var fieldsToRename = fields.Where(f => !f.BackingIdentifier.Equals(f.Declaration.IdentifierName));

            foreach (var field in fieldsToRename)
            {
                if (!(field.Declaration.Context is IIdentifierContext context))
                {
                    throw new ArgumentException();
                }

                rewriter.Replace(context.IdentifierTokens, field.BackingIdentifier);
            }
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
