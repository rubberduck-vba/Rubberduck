using System;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ISafeComWrapper : INullObjectWrapper, IDisposable
    {
        void Release(bool final = false);
        bool HasBeenReleased { get; }
    }

    public interface ISafeComWrapper<out T> : ISafeComWrapper
    {
        new T Target { get; }
    }

    public interface INullObjectWrapper
    {
        object Target { get; }
        bool IsWrappingNullReference { get; }
    }
}