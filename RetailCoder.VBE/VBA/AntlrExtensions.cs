using System.Collections.Generic;
using Antlr4.Runtime.Tree;
using Rubberduck.VBA.Grammar;
using Rubberduck.VBA.ParseTreeListeners;

namespace Rubberduck.VBA
{
    /// <summary>
    /// Extension methods for <c>IParseTree</c> and <c>ParserRuleContext</c>.
    /// </summary>
    public static class AntlrExtensions
    {
        public static IEnumerable<TContext> GetContexts<TListener, TContext>(this IParseTree parseTree, TListener listener)
            where TListener : IExtensionListener<TContext>, IParseTreeListener
            where TContext : class
        {
            try
            {
                var walker = new ParseTreeWalker();
                walker.Walk(listener, parseTree);
            }
            catch (WalkerCancelledException)
            {
                // swallow. this exception is thrown when a listener doesn't need to walk an entire parse tree.
            }

            return listener.Members;
        }

        /// <summary>
        /// Assuming the specified identifier is in scope,
        /// returns <c>true</c> if its name matches that of the used variable.
        /// </summary>
        /// <returns></returns>
        public static bool IsIdentifierUsage(this VisualBasic6Parser.VariableCallStmtContext usage, VisualBasic6Parser.AmbiguousIdentifierContext identifier)
        {
            return usage.ambiguousIdentifier().GetText() == identifier.GetText();
        }
    }
}