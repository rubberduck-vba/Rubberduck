using System.Runtime.CompilerServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public abstract partial class ComSafeBase : IComSafe
    {
        public abstract void Add(ISafeComWrapper comWrapper);

        public abstract bool TryRemove(ISafeComWrapper comWrapper);

        //We do not use GetHashCode because subclasses of SafeComWrapper<T> overwrite this method 
        //and we need to distinguish between individual instances.
        protected int GetComWrapperObjectHashCode(ISafeComWrapper comWrapper)
        {
            return RuntimeHelpers.GetHashCode(comWrapper);
        }

        private bool _disposed;
        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            TraceDispose();
        }
    }
}