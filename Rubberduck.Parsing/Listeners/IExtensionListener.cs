using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Listeners
{
    public interface IExtensionListener<TContext>
        where TContext : ParserRuleContext
    {
        IEnumerable<QualifiedContext<TContext>> Members { get; }
    }
}