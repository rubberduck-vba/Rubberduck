using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public interface IComSafe: IDisposable
    {
        void Add(ISafeComWrapper comWrapper);
        bool TryRemove(ISafeComWrapper comWrapper);
    }
}
