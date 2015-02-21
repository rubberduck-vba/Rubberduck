using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public interface IExtensionListener<TContext>
        where TContext : class
    {
        IEnumerable<QualifiedContext<TContext>> Members { get; }
    }
}