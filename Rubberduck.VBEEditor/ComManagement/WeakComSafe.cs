using System;
using System.Collections.Concurrent;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;
using System.Collections.Generic;

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
                _comWrapperCache.AddOrUpdate(
                    GetComWrapperObjectHashCode(comWrapper), 
                    key =>
                    {
                        TraceAdd(comWrapper);
                        return (DateTime.UtcNow, new WeakReference<ISafeComWrapper>(comWrapper));
                    },
                    (key, value) =>
                    {
                        TraceUpdate(comWrapper);
                        return (value.insertTime, new WeakReference<ISafeComWrapper>(comWrapper));
                    });
            }

        }

        public override bool TryRemove(ISafeComWrapper comWrapper)
        {
            if (_disposed || comWrapper == null)
            {
                return false;
            }

            var result = _comWrapperCache.TryRemove(GetComWrapperObjectHashCode(comWrapper), out _);
            TraceRemove(comWrapper, result);

            return result;
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            base.Dispose(disposing);

            foreach (var weakReference in _comWrapperCache.Values)
            {
                if (weakReference.weakRef.TryGetTarget(out var comWrapper))
                {
                    comWrapper.Dispose();
                }
            }

            _comWrapperCache.Clear();
        }

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
    }
}

