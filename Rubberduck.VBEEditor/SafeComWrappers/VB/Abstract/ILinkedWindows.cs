using System;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.Abstract
{
    public interface ILinkedWindows : ISafeComWrapper, IComCollection<IWindow>, IEquatable<ILinkedWindows>
    {
        IVBE VBE { get; }
        IWindow Parent { get; }
        void Remove(IWindow window);
        void Add(IWindow window);
    }
}