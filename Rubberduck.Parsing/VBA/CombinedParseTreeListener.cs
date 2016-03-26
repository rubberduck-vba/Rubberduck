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
        private List<IParseTreeListener> _listeners;
        public CombinedParseTreeListener(IParseTreeListener[] listeners)
        {
            _listeners = listeners.ToList();
        }

        public void EnterEveryRule(ParserRuleContext ctx)
        {
            _listeners.ForEach(l => l.EnterEveryRule(ctx));
        }

        public void ExitEveryRule(ParserRuleContext ctx)
        {
            _listeners.ForEach(l => l.ExitEveryRule(ctx));
        }

        public void VisitErrorNode(IErrorNode node)
        {
            _listeners.ForEach(l => l.VisitErrorNode(node));
        }

        public void VisitTerminal(ITerminalNode node)
        {
            _listeners.ForEach(l => l.VisitTerminal(node));
        }
    }
}
