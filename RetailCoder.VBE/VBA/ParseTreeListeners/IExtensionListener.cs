using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public interface IExtensionListener<out TContext>
        where TContext : class
    {
        IEnumerable<TContext> Members { get; }
    }
}