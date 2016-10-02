using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public abstract class SafeComWrapper<T> : IDisposable
        where T : class 
    {
        private readonly T _comObject;
        private bool _isDisposed;

        protected SafeComWrapper(T comObject)
        {
            _comObject = comObject;
        }

        protected internal T ComObject
        {
            get
            {
                ThrowIfDisposed();
                return _comObject;
            }
        }

        public bool IsNull
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

        ~SafeComWrapper()
        {
            Dispose();
        }
    }
}