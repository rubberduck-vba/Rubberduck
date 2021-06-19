using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public interface IDeleteDeclarationModifyEndOfStatementContentModel
    {
        bool RemoveDeclarationLogicalLineComment { set;  get; }
        bool InsertValidationTODOForRetainedComments { set;  get; }
    }

    public class DeleteDeclarationEndOfStatementContentModifier : IDeleteDeclarationEndOfStatementContentModifier
    {
        private readonly IEOSContextContentProviderFactory _eosContextContentProviderFactory;
        public DeleteDeclarationEndOfStatementContentModifier(IEOSContextContentProviderFactory eosContextContentProviderFactory)
        {
            _eosContextContentProviderFactory = eosContextContentProviderFactory;
        }

        public void ModifyEndOfStatementContextContent(IDeclarationDeletionTarget deleteTarget, IDeleteDeclarationModifyEndOfStatementContentModel model, IModuleRewriter rewriter)
        {
            var precedingEOS = _eosContextContentProviderFactory.Create(deleteTarget.PrecedingEOSContext, rewriter);
            var deletedDeclarationEOS = _eosContextContentProviderFactory.Create(deleteTarget.EndOfStatementContext, rewriter);

            var contextTargetForMergedContent = precedingEOS.EOSContext ?? deletedDeclarationEOS.EOSContext;

            if (contextTargetForMergedContent is null || deletedDeclarationEOS.IsNullEOS)
            {
                //No EOSContexts to use
                return;
            }

            if (deletedDeclarationEOS.EOSContext.GetText().StartsWith(": "))
            {
                rewriter.Remove(deletedDeclarationEOS.EOSContext);
                return;
            }

            RemoveDeclarationLogicalLineComment(deletedDeclarationEOS, model, rewriter);

            InjectTODOForRetainedComments(precedingEOS, model, rewriter);

            if (contextTargetForMergedContent == deletedDeclarationEOS.EOSContext)
            {
                rewriter.Replace(contextTargetForMergedContent, string.Empty);
            }

            else if (deleteTarget.HasPrecedingLabel(out var labelContext))
            {
                ModifyDeclarationWithLineLabel(deleteTarget, precedingEOS, deletedDeclarationEOS, labelContext, rewriter);
            }

            else if (deletedDeclarationEOS.ModifiedContentContainsCommentMarker)
            {
                ModifyEOSContextWithRetainedDeleteTargetComments(precedingEOS, deletedDeclarationEOS, contextTargetForMergedContent, rewriter);
            }
            else
            {
                var replacementContent = $"{precedingEOS.ContentPriorToSeparationAndIndentation}{precedingEOS.Separation}{deletedDeclarationEOS.Indentation}";

                //Hack: if the deleted declaration is a Member (it's in the Module Code Section), then ensure there is minimally 
                //2 newlines at the end of the newly generated replacement content.
                //TODO: Make IndenterSettings available here to set the minimum separation value
                if (deleteTarget.PrecedingEOSContext.TryGetPrecedingContext<VBAParser.ModuleBodyElementContext>(out _))
                {
                    var minSeparationForProceduresAfterDeletions = string.Concat(Enumerable.Repeat(Environment.NewLine, 2));
                    if (!precedingEOS.Separation.Contains(minSeparationForProceduresAfterDeletions))
                    {
                        replacementContent = $"{precedingEOS.ContentPriorToSeparationAndIndentation}{precedingEOS.Separation}{Environment.NewLine}{deletedDeclarationEOS.Indentation}";
                    }
                }

                ModifyEOSContext(replacementContent, deletedDeclarationEOS, contextTargetForMergedContent, rewriter);
            }
        }

        private static void ModifyEOSContextWithRetainedDeleteTargetComments(IEOSContextContentProvider precedingEOS, IEOSContextContentProvider deletedDeclarationEOS, ParserRuleContext targetContext, IModuleRewriter rewriter)
        {
            var replacementContentWithCommentsRetained = precedingEOS.ModifiedContentContainsCommentMarker
                //Both the declarationEOS and the precedingEOS have comments - concatenate them but leave out the precedingEOS Separation and Indentation
                ? $"{precedingEOS.ContentPriorToSeparationAndIndentation}{deletedDeclarationEOS.ModifiedEOSContent}"
                //Only the declarationEOS has comments, the precedingEOS is entirely whitespace and newlines
                : $"{precedingEOS.Separation}{deletedDeclarationEOS.ContentFreeOfStartingNewLines}";

            ModifyEOSContext(replacementContentWithCommentsRetained, deletedDeclarationEOS, targetContext, rewriter);
        }

        private static void ModifyEOSContext(string replacementContent, IEOSContextContentProvider deletedDeclarationEOS, ParserRuleContext contextTarget, IModuleRewriter rewriter)
        {
            if (!deletedDeclarationEOS.IsNullEOS)
            {
                rewriter.Remove(deletedDeclarationEOS.EOSContext);
            }

            rewriter.Replace(contextTarget, replacementContent);
        }

        private static void RemoveDeclarationLogicalLineComment(IEOSContextContentProvider deletedDeclarationEOS, IDeleteDeclarationModifyEndOfStatementContentModel model, IModuleRewriter rewriter)
        {
            if (model.RemoveDeclarationLogicalLineComment && deletedDeclarationEOS.HasDeclarationLogicalLineComment)
            {
                rewriter.Remove(deletedDeclarationEOS.DeclarationLogicalLineCommentContext.children.OfType<ParserRuleContext>().First());
            }
        }


        private static void InjectTODOForRetainedComments(IEOSContextContentProvider precedingEOS, IDeleteDeclarationModifyEndOfStatementContentModel model, IModuleRewriter rewriter)
        {
            //TODO: this only deals with comments in the preceding context.  If there are comments in the target context - shouldn't
            //they also get injected with the TODO message?
            if (!model.InsertValidationTODOForRetainedComments || precedingEOS.EOSContext is null)
            {
                return;
            }

            var injectedTODOContent = Resources.Refactorings.Refactorings.ImplementInterface_TODO;

            foreach (var comment in precedingEOS.AllComments.Where(c => c != precedingEOS.DeclarationLogicalLineCommentContext.GetDescendent<VBAParser.CommentContext>()))
            {
                var content = comment.GetText();
                var indexOfFirstCommentMarker = content.IndexOf(Tokens.CommentMarker);
                var newContent = $"{content.Substring(0, indexOfFirstCommentMarker + 1)}{injectedTODOContent}{content.Substring(indexOfFirstCommentMarker + 1)}";
                rewriter.Replace(comment, newContent);
            }
        }

        private static void ModifyDeclarationWithLineLabel(IDeclarationDeletionTarget deleteTarget, IEOSContextContentProvider precedingEOS, IEOSContextContentProvider deletedDeclarationEOS, VBAParser.StatementLabelDefinitionContext labelContext, IModuleRewriter rewriter)
        {
            if (deletedDeclarationEOS.ModifiedContentContainsCommentMarker)
            {
                var replacementContent = precedingEOS.ModifiedContentContainsCommentMarker
                //Both the declarationEOS and the precedingEOS have comments - concatenate them leaving out the precedingEOS Separation and Indentation
                ? $"{precedingEOS.ContentPriorToSeparationAndIndentation}{labelContext.GetText()}{deletedDeclarationEOS.ModifiedEOSContent}"
                //Only the declarationEOS has comments, the precedingEOS is entirely whitespace and newlines
                : $"{precedingEOS.Separation}{labelContext.GetText()}{deletedDeclarationEOS.ModifiedEOSContent}";

                var targetContext = precedingEOS.EOSContext ?? deletedDeclarationEOS.EOSContext;
                ModifyEOSContext(replacementContent, deletedDeclarationEOS, targetContext as ParserRuleContext, rewriter);
                return;
            }

            //Deletes only the declaration and not the label.  Appends the EOS content to the remaining
            //content of the declaration line (the label) and replaces the target context.  The prior EOS 
            //is injected as part of the TargetContext replacement
            var blockStmt = deleteTarget.DeleteContext.GetAncestor<VBAParser.BlockStmtContext>();
            var modifedContent = GetModifiedContextText(blockStmt, rewriter);
                    
            var deleteTargetReplacement = $"{modifedContent}{deletedDeclarationEOS.Separation}{deletedDeclarationEOS.Indentation}";
            rewriter.Replace(deleteTarget.TargetContext, deleteTargetReplacement);

            var eosReplacementContent = $"{precedingEOS.ContentPriorToSeparationAndIndentation}{precedingEOS.Separation}";
            rewriter.Replace(precedingEOS.EOSContext, eosReplacementContent);
            rewriter.Remove(deletedDeclarationEOS.EOSContext);
        }

        private static string GetModifiedContextText(ParserRuleContext prContext, IModuleRewriter rewriter)
            => rewriter.GetText(prContext.Start.TokenIndex, prContext.Stop.TokenIndex);

    }
}
