using System.Collections.Generic;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IComCollection<out TItem> : IEnumerable<TItem>
    {
        int Count { get; }
        TItem this[object index] { get; }
    }
}