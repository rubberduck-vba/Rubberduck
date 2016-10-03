using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public abstract class SafeComWrapper<T> : IDisposable, IEquatable<SafeComWrapper<T>> 
        where T : class 
    {
        private readonly T _comObject;
        private bool _isDisposed;

        protected SafeComWrapper(T comObject)
        {
            _comObject = comObject;
        }

        public T ComObject
        {
            get
            {
                ThrowIfDisposed();
                return _comObject;
            }
        }

        public bool IsWrappingNullReference
        {
            get
            {
                ThrowIfDisposed();
                return _comObject == null;
            }
        }

        protected TResult InvokeResult<TResult>(Func<TResult> member)
        {
            ThrowIfDisposed();
            try
            {
                return member.Invoke();
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected void Invoke(Action member)
        {
            ThrowIfDisposed();
            try
            {
                member.Invoke();
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected void ThrowIfDisposed()
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException("Object has been disposed.");
            }
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            if (_comObject != null)
            {
                Marshal.ReleaseComObject(_comObject);
            }

            _isDisposed = true;
        }

        public bool Equals(SafeComWrapper<T> other)
        {
            ThrowIfDisposed();
            if (IsWrappingNullReference)
            {
                return other.IsWrappingNullReference;
            }

            return ReferenceEquals(_comObject, other._comObject);
        }

        public override bool Equals(object obj)
        {
            ThrowIfDisposed();
            return obj is T && ReferenceEquals(_comObject, obj); // bug: COM object isn't reliable for reference equality
        }

        public override int GetHashCode()
        {
            ThrowIfDisposed();
            return _comObject.GetHashCode(); // bug: COM object isn't reliable for a hashcode
        }

        ~SafeComWrapper()
        {
            Dispose();
        }
    }
}