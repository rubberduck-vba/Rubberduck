using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.DisposableWrappers
{
    public abstract class WrapperBase<T> : IDisposable
        where T : class 
    {
        private readonly T _item;
        private bool _isDisposed;

        protected WrapperBase(T item)
        {
            _item = item;
        }

        protected internal T Item
        {
            get
            {
                ThrowIfDisposed(_isDisposed);
                return _item;
            }
        }

        public bool IsNull
        {
            get
            {
                ThrowIfDisposed(_isDisposed);
                return _item == null;
            }
        }

        public static bool operator ==(WrapperBase<T> object1, WrapperBase<T> object2)
        {
            if (object1 != null && object1.IsNull)
            {
                return (object)object2 == null;
            }
            if (object2 != null && object2.IsNull)
            {
                return (object)object1 == null;
            }

            return ReferenceEquals(object1, object2);
        }

        public static bool operator !=(WrapperBase<T> object1, WrapperBase<T> object2)
        {
            return !(object1 == object2);
        }

        protected static TResult InvokeMemberValue<TResult>(Func<TResult> member)
        {
            try
            {
                return member.Invoke();
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static TResult InvokeMemberValue<T, TResult>(Func<T, TResult> member, T param)
        {
            try
            {
                return member.Invoke(param);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static TResult InvokeMemberValue<T1, T2, TResult>(Func<T1, T2, TResult> member, T1 param1, T2 param2)
        {
            try
            {
                return member.Invoke(param1, param2);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static void InvokeMember(Action member)
        {
            try
            {
                member.Invoke();
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static void InvokeMember<T>(Action<T> member, T param)
        {
            try
            {
                member.Invoke(param);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static void InvokeMember<T1, T2>(Action<T1, T2> member, T1 param1, T2 param2)
        {
            try
            {
                member.Invoke(param1, param2);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static void InvokeMember<T1, T2, T3>(Action<T1, T2, T3> member, T1 param1, T2 param2, T3 param3)
        {
            try
            {
                member.Invoke(param1, param2, param3);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static void InvokeMember<T1, T2, T3, T4>(Action<T1, T2, T3, T4> member, T1 param1, T2 param2, T3 param3, T4 param4)
        {
            try
            {
                member.Invoke(param1, param2, param3, param4);
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        protected static void ThrowIfDisposed(bool isDisposed)
        {
            if (isDisposed) { throw new ObjectDisposedException("Object has been disposed."); }
        }

        protected void ThrowIfDisposed()
        {
            ThrowIfDisposed(_isDisposed);
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            Marshal.ReleaseComObject(_item);
            _isDisposed = true;
        }
    }
}