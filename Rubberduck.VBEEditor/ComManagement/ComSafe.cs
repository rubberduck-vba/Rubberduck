using System;
using System.Collections.Concurrent;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public class ComSafe: IComSafe
    {
        //We override the equality comparison and hash code because subclasses of SafeComWrapper<T> override the corresponding methods.
        //We need to distinguish between the individual wrapper instances no matter whether they are semantically equal.
        private readonly ConcurrentDictionary<ISafeComWrapper, byte> _comWrapperCache = new ConcurrentDictionary<ISafeComWrapper, byte>(new ReferenceEqualityComparer());


        public void Add(ISafeComWrapper comWrapper)
        {
            if (comWrapper != null)
            {
                _comWrapperCache.AddOrUpdate(comWrapper, key => 1, (key, value) => value);
            }
            
        }

        public bool TryRemove(ISafeComWrapper comWrapper)
        {
            return !_disposed && comWrapper != null && _comWrapperCache.TryRemove(comWrapper, out _);
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            foreach (var comWrapper in _comWrapperCache.Keys)
            {
                comWrapper.Dispose();
            }

            _comWrapperCache.Clear();
        }
    }

    public static class ComSafeManager
    {
        private static Lazy<ComSafe> _comSafe = new Lazy<ComSafe>();

        public static IComSafe GetCurrentComSafe()
        {
            return _comSafe.Value;
        }

        public static void ResetComSafe()
        {
            _comSafe = new Lazy<ComSafe>();
        }
    }
}
