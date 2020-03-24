using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField.Extensions
{
    public static class IModuleRewriterExtensions
    {
        public static string GetText(this IModuleRewriter rewriter, int maxConsecutiveNewLines)
        {
            var result = rewriter.GetText();
            var target = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            var replacement = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines - 1).ToList());
            for (var counter = 1; counter < 10 && result.Contains(target); counter++)
            {
                result = result.Replace(target, replacement);
            }
            return result;
        }

        public static void InsertAtEndOfFile(this IModuleRewriter rewriter, string content)
        {
            if (content == string.Empty) { return; }

            rewriter.InsertBefore(rewriter.TokenStream.Size - 1, content);
        }

        public static void MakeImplicitDeclarationTypeExplicit(this IModuleRewriter rewriter, Declaration element)
        {
            if (!element.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
            {
                rewriter.InsertAfter(element.Context.Stop.TokenIndex, $" {Tokens.As} {element.AsTypeName}");
            }
        }

        public static void Rename(this IModuleRewriter rewriter, Declaration target, string newName)
        {
            if (!(target.Context is IIdentifierContext context))
            {
                throw new ArgumentException();
            }

            rewriter.Replace(context.IdentifierTokens, newName);
        }

        public static void SetVariableVisiblity(this IModuleRewriter rewriter, Declaration element, string visibility)
        {
            if (!element.IsVariable()) { throw new ArgumentException(); }

            var variableStmtContext = element.Context.GetAncestor<VBAParser.VariableStmtContext>();
            var visibilityContext = variableStmtContext.GetChild<VBAParser.VisibilityContext>();

            if (visibilityContext != null)
            {
                rewriter.Replace(visibilityContext, visibility);
                return;
            }
            rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{visibility} ");
        }
    }
}
