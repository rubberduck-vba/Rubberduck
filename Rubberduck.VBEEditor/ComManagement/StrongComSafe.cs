using System.Collections.Concurrent;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public class StrongComSafe: ComSafeBase
    {
        //We override the equality comparison and hash code because subclasses of SafeComWrapper<T> override the corresponding methods.
        //We need to distinguish between the individual wrapper instances no matter whether they are semantically equal.
        private readonly ConcurrentDictionary<int, ISafeComWrapper> _comWrapperCache = new ConcurrentDictionary<int, ISafeComWrapper>();

        public override void Add(ISafeComWrapper comWrapper)
        {
            if (comWrapper != null)
            {
#if DEBUG
                Trace = GetStackTrace(3, 3);
#endif
                _comWrapperCache.AddOrUpdate(
                    GetComWrapperObjectHashCode(comWrapper), 
                    value => comWrapper, 
                    (key, value) =>
                    {
#if DEBUG
                        System.Diagnostics.Debug.Assert(false);
#endif
                        return value;
                    });
            }
        }

        public override bool TryRemove(ISafeComWrapper comWrapper)
        {
            return !_disposed && comWrapper != null &&
                   _comWrapperCache.TryRemove(GetComWrapperObjectHashCode(comWrapper), out _);
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            foreach (var comWrapper in _comWrapperCache.Values)
            {
                comWrapper.Dispose();
            }

            _comWrapperCache.Clear();
        }

#if DEBUG
        protected override IDictionary<int, ISafeComWrapper> GetWrappers()
        {
            return _comWrapperCache;
        }
#endif
    }
}
