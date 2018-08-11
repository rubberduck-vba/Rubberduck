using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IEvents : ISafeComWrapper, IEquatable<IEvents>
    {
        ICommandBarEvents CommandBarEvents { get; }
    }
}
