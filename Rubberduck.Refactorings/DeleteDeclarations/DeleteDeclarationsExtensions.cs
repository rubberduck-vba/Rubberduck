using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal static class ParserRuleContextExtensions
    {
        public static IEnumerable<ParserRuleContext> GetChildrenOfType<T>(this ParserRuleContext parent) where T : ParserRuleContext
            => parent?.children?.OfType<T>().Cast<ParserRuleContext>() ?? Enumerable.Empty<ParserRuleContext>();

        public static string CurrentContent(this ParserRuleContext context, IModuleRewriter rewriter)
            => context != null ? rewriter.GetText(context.Start.TokenIndex, context.Stop.TokenIndex) : string.Empty;

        public static VBAParser.EndOfStatementContext GetFollowingEndOfStatementContext(this ParserRuleContext context)
        {
            context.TryGetFollowingContext(out VBAParser.EndOfStatementContext eos);
            return eos;
        }

        public static VBAParser.EndOfStatementContext GetPrecedingEndOfStatementContext(this ParserRuleContext context)
        {
            context.TryGetPrecedingContext(out VBAParser.EndOfStatementContext eos);
            return eos;
        }
    }

    internal static class EndOfStatementContextExtensions
    {
        private const string EndOfStatementColon = ": ";

        public static string GetSeparation(this VBAParser.EndOfStatementContext eosContext)
            => string.Concat(eosContext.GetSeparationAndIndentationContent().TakeWhile(c => c != ' '));

        public static string GetIndentation(this VBAParser.EndOfStatementContext eosContext)
            => string.Concat(eosContext.GetSeparationAndIndentationContent().SkipWhile(c => c == '\r' || c == '\n'));

        public static string GetSeparationAndIndentationContent(this VBAParser.EndOfStatementContext eosContext)
        {
            var eosContent = eosContext?.GetText() ?? string.Empty;
            if ((eosContent?.Length ?? 0) <= 0)
            {
                return string.Empty;
            }

            return eosContent.StartsWith(EndOfStatementColon)
              ? string.Empty
              : Regex.Match(eosContent, @"(\r\n)+\s*$").Value;
        }

        public static IEnumerable<VBAParser.CommentContext> GetAllComments(this VBAParser.EndOfStatementContext eosContext)
            => eosContext?.children.OfType<VBAParser.IndividualNonEOFEndOfStatementContext>()
                .Select(child => child.GetDescendent<VBAParser.CommentContext>())
                .Where(ch => ch != null)
            ?? Enumerable.Empty<VBAParser.CommentContext>();

    }
}
