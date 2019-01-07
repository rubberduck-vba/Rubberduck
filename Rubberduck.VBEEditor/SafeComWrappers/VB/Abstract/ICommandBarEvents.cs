using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICommandBarEvents : ISafeComWrapper, IComIndexedProperty<ICommandBarButtonEvents>, IEquatable<ICommandBarEvents>
    {
    }
}
