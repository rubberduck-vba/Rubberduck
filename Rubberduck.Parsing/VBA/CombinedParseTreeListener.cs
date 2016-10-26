using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// A Class combining an arbitrary number of IParseTreeListener instances into one single instance
    /// </summary>
    public class CombinedParseTreeListener : IParseTreeListener
    {
        private readonly IReadOnlyList<IParseTreeListener> _listeners;
        public CombinedParseTreeListener(IEnumerable<IParseTreeListener> listeners)
        {
            _listeners = listeners.Where(listener => listener != null).ToList();
        }

        public void EnterEveryRule(ParserRuleContext ctx)
        {
            foreach (var listener in _listeners)
            {
                listener.EnterEveryRule(ctx);
                ctx.EnterRule(listener);
            }
        }

        public void ExitEveryRule(ParserRuleContext ctx)
        {
            foreach (var listener in _listeners)
            {
                listener.ExitEveryRule(ctx);
                ctx.ExitRule(listener);
            }
        }

        public void VisitErrorNode(IErrorNode node)
        {
            foreach (var listener in _listeners)
            {
                listener.VisitErrorNode(node);
            }
        }

        public void VisitTerminal(ITerminalNode node)
        {
            foreach (var listener in _listeners)
            {
                listener.VisitTerminal(node);
            }
        }
    }
}
