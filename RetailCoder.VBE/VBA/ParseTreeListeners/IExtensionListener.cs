using System.Collections.Generic;

namespace Rubberduck.VBA.ParseTreeListeners
{
    public interface IExtensionListener<TContext>
        where TContext : class
    {
        IEnumerable<QualifiedContext<TContext>> Members { get; }
    }
}