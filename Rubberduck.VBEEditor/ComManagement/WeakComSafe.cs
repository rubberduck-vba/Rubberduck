using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

#if DEBUG
using System.Diagnostics;
using System.Linq;
#endif

namespace Rubberduck.VBEditor.ComManagement
{
    public class WeakComSafe : ComSafeBase
    {
        //We use weak references to allow the GC to reclaim RCWs earlier if possible.
        private readonly ConcurrentDictionary<int, (DateTime insertTime, WeakReference<ISafeComWrapper> weakRef)> _comWrapperCache = new ConcurrentDictionary<int, (DateTime, WeakReference<ISafeComWrapper>)>();

        public override void Add(ISafeComWrapper comWrapper)
        {
            if (comWrapper != null)
            {
#if DEBUG
                Trace = GetStackTrace(3, 3);
#endif
                _comWrapperCache.AddOrUpdate(
                    GetComWrapperObjectHashCode(comWrapper), 
                    key => (DateTime.UtcNow, new WeakReference<ISafeComWrapper>(comWrapper)),
                    (key, value) =>
                    {
#if DEBUG
                        Debug.Assert(false);
#endif
                        return (value.insertTime, new WeakReference<ISafeComWrapper>(comWrapper));
                    });
            }

        }

        public override bool TryRemove(ISafeComWrapper comWrapper)
        {
            return !_disposed && comWrapper != null && _comWrapperCache.TryRemove(GetComWrapperObjectHashCode(comWrapper), out _);
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            foreach (var weakReference in _comWrapperCache.Values)
            {
                if (weakReference.weakRef.TryGetTarget(out var comWrapper))
                {
                    comWrapper.Dispose();
                }
            }

            _comWrapperCache.Clear();
        }

#if DEBUG
        protected override IDictionary<int, ISafeComWrapper> GetWrappers()
        {
            var dictionary = new Dictionary<int, ISafeComWrapper>();
            foreach (var kvp in _comWrapperCache.OrderBy(kvp => kvp.Value.insertTime))
            {
                if (!kvp.Value.weakRef.TryGetTarget(out var target))
                {
                    target = null;
                }
                dictionary.Add(kvp.Key, target);   
            }

            return dictionary;
        }
#endif
    }
}
