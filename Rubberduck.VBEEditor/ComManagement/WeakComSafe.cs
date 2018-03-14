using System;
using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public class WeakComSafe : IComSafe
    {
        //We use weak references to allow the GC to reclaim RCWs earlier if possible.
        private readonly ConcurrentDictionary<int, WeakReference<ISafeComWrapper>> _comWrapperCache = new ConcurrentDictionary<int, WeakReference<ISafeComWrapper>>();


        public void Add(ISafeComWrapper comWrapper)
        {
            if (comWrapper != null)
            {
                _comWrapperCache.AddOrUpdate(
                    GetComWrapperObjectHashCode(comWrapper), 
                    key => new WeakReference<ISafeComWrapper>(comWrapper), 
                    (key, value) => new WeakReference<ISafeComWrapper>(comWrapper));
            }

        }

        //We do not use GetHashCode because subclasses of SafeComWrapper<T> overwrite this method 
        //and we need to distinguish between individual instances.
        private int GetComWrapperObjectHashCode(ISafeComWrapper comWrapper)
        {
            return RuntimeHelpers.GetHashCode(comWrapper);
        }

        public bool TryRemove(ISafeComWrapper comWrapper)
        {
            return !_disposed && comWrapper != null && _comWrapperCache.TryRemove(GetComWrapperObjectHashCode(comWrapper), out _);
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            foreach (var weakReference in _comWrapperCache.Values)
            {
                if(weakReference.TryGetTarget(out var comWrapper))
                {
                    comWrapper.Dispose();
                }
            }

            _comWrapperCache.Clear();
        }
    }
}
