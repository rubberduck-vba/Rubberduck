using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public class SafeIDispatchWrapper<TDispatch> : SafeIDispatchWrapper
    {
        public SafeIDispatchWrapper(TDispatch target, bool rewrapping = false) : base(target, rewrapping)
        { }

        public new TDispatch Target => (TDispatch) base.Target;
    }

    public class SafeIDispatchWrapper : SafeComWrapper<dynamic>
    {
        public SafeIDispatchWrapper(object target, bool rewrapping = false) : base(target, rewrapping)
        {
            if (!Marshal.IsComObject(target))
            {
                throw new ArgumentException("The target object must be a COM object");
            }

            IDispatchPointer = GetPointer(target);
        }

        public IntPtr IDispatchPointer { get; }

        public void Invoke(Action<dynamic> action)
        {
            action.Invoke(Target);
        }

        public override bool Equals(ISafeComWrapper<dynamic> other)
        {
            if (other.IsWrappingNullReference || IsWrappingNullReference)
            {
                return false;
            }

            return IDispatchPointer == GetPointer(other);
        }

        public override int GetHashCode()
        {
            return unchecked((int)IDispatchPointer); 
        }

        private static IntPtr GetPointer(object target)
        {
            var result = IntPtr.Zero;
            try
            {
                result = Marshal.GetIDispatchForObject(target);
            }
            finally
            {
                if (result != IntPtr.Zero)
                {
                    Marshal.Release(result);
                }
            }

            return result;
        }
    }
}
