using System;
using Rubberduck.VBEditor.SafeComWrappers.VBA.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract
{
    public interface ILinkedWindows : ISafeComWrapper, IComCollection<IWindow>, IEquatable<ILinkedWindows>
    {
        IVBE VBE { get; }
        IWindow Parent { get; }
        void Remove(IWindow window);
        void Add(IWindow window);
    }
}