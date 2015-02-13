using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public interface IExtensionListener<out TContext>
        where TContext : ParserRuleContext
    {
        IEnumerable<TContext> Members { get; }
    }
}