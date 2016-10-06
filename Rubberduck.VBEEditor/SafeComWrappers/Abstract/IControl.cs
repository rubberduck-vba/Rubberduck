using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface IControl : ISafeComWrapper, IEquatable<IControl>
    {
        string Name { get; set; }
    }
}