using System.Collections.Generic;

namespace Rubberduck.Parsing.Listeners
{
    public interface IExtensionListener<TContext>
        where TContext : class
    {
        IEnumerable<QualifiedContext<TContext>> Members { get; }
    }
}