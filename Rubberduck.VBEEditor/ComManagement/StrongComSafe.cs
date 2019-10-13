using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public class StrongComSafe: ComSafeBase
    {
        //We override the equality comparison and hash code because subclasses of SafeComWrapper<T> override the corresponding methods.
        //We need to distinguish between the individual wrapper instances no matter whether they are semantically equal.
        private readonly ConcurrentDictionary<ISafeComWrapper, byte> _comWrapperCache = new ConcurrentDictionary<ISafeComWrapper, byte>(new ReferenceEqualityComparer());

        public override void Add(ISafeComWrapper comWrapper)
        {
            if (comWrapper != null)
            {
                _comWrapperCache.AddOrUpdate(
                    comWrapper, 
                    key =>
                    {
                        TraceAdd(comWrapper);
                        return 1;
                    }, 
                    (key, value) =>
                    {
                        TraceUpdate(comWrapper);
                        return value;
                    });
            }
        }

        public override bool TryRemove(ISafeComWrapper comWrapper)
        {
            if (_disposed || comWrapper == null)
            {
                return false;
            }

            var result = _comWrapperCache.TryRemove(comWrapper, out _);
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

            foreach (var comWrapper in _comWrapperCache.Keys)
            {
                comWrapper.Dispose();
            }

            _comWrapperCache.Clear();
        }

        protected override IDictionary<int, ISafeComWrapper> GetWrappers()
        {
            return _comWrapperCache.Keys.ToDictionary(GetComWrapperObjectHashCode);
        }
    }
}

