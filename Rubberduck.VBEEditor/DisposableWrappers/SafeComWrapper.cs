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

        #region protected TResult InvokeResult<TResult>
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

        protected TResult InvokeResult<T, TResult>(Func<T, TResult> member, T param)
        {
            ThrowIfDisposed();
            try
            {
                return member.Invoke(param);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected TResult InvokeResult<T1, T2, TResult>(Func<T1, T2, TResult> member, T1 param1, T2 param2)
        {
            ThrowIfDisposed();
            try
            {
                return member.Invoke(param1, param2);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected TResult InvokeResult<T1, T2, T3, TResult>(Func<T1, T2, T3, TResult> member, T1 param1, T2 param2, T3 param3)
        {
            ThrowIfDisposed();
            try
            {
                return member.Invoke(param1, param2, param3);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected TResult InvokeResult<T1, T2, T3, T4, TResult>(Func<T1, T2, T3, T4, TResult> member, T1 param1, T2 param2, T3 param3, T4 param4)
        {
            ThrowIfDisposed();
            try
            {
                return member.Invoke(param1, param2, param3, param4);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }
        #endregion

        #region protected void Invoke
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

        protected void Invoke<T>(Action<T> member, T param)
        {
            ThrowIfDisposed();
            try
            {
                member.Invoke(param);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected void Invoke<T1, T2>(Action<T1, T2> member, T1 param1, T2 param2)
        {
            ThrowIfDisposed();
            try
            {
                member.Invoke(param1, param2);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected void Invoke<T1, T2, T3>(Action<T1, T2, T3> member, T1 param1, T2 param2, T3 param3)
        {
            ThrowIfDisposed();
            try
            {
                member.Invoke(param1, param2, param3);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected void Invoke<T1, T2, T3, T4>(Action<T1, T2, T3, T4> member, T1 param1, T2 param2, T3 param3, T4 param4)
        {
            ThrowIfDisposed();
            try
            {
                member.Invoke(param1, param2, param3, param4);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected void Invoke<T1, T2, T3, T4, T5>(Action<T1, T2, T3, T4, T5> member, T1 param1, T2 param2, T3 param3, T4 param4, T5 param5)
        {
            ThrowIfDisposed();
            try
            {
                member.Invoke(param1, param2, param3, param4, param5);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }
        #endregion

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