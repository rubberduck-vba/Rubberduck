using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public interface IExtensionListener<out TContext>
    {
        IEnumerable<TContext> Members { get; }
    }
}