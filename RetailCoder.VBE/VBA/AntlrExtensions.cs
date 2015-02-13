using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
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
            where TContext : ParserRuleContext
        {
            var walker = new ParseTreeWalker();
            walker.Walk(listener, parseTree);

            return listener.Members;
        }

    }
}