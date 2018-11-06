using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    /// <summary>
    /// Creates a new IDispatch-based wrapper for a provided
    /// COM interface. This is internal and should be used only
    /// within the Rubberduck.VBEditor.* projects.
    /// </summary>
    /// <remarks>
    /// To avoid exposing additional interop libraries to other projects
    /// and violating the separation, only the non-generic version of
    /// <see cref="SafeIDispatchWrapper" /> be provided for consumption outside
    /// the Rubberduck.VBEditor.* projects.
    /// Within those projects, the class is useful for wrapping COM interfaces
    /// that do not have a corresponding  <see cref="ISafeComWrapper{T}" />
    /// implementations to ensure those are managed and will not leak.
    /// </remarks>
    /// <typeparam name="TDispatch">COM interface to wrap</typeparam>
    /// <inheritdoc />
    internal class SafeIDispatchWrapper<TDispatch> : SafeIDispatchWrapper
    {
        internal SafeIDispatchWrapper(TDispatch target, bool rewrapping = false) : base(target, rewrapping)
        { }

        public new TDispatch Target => (TDispatch) base.Target;
    }

    /// <summary>
    /// Provide a IDispatch-based (e.g. late-bound) access to a COM object.
    /// Use <see cref="Invoke"/> to work with the object. The
    /// <see cref="SafeIDispatchWrapper.Target"/> must be a raw COM object without
    /// any wrappers. 
    /// </summary>
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

        // ReSharper disable once InconsistentNaming
        public IntPtr IDispatchPointer { get; }

        /// <summary>
        /// Use the method to encapsulate operations against the late-bound COM object.
        /// It is caller's responsibilty to handle any exceptions that may result as part
        /// of the operation. 
        /// </summary>
        /// <param name="action">
        /// A method that perform work against the late-bound COM object provided as the
        /// parameter to the function
        /// </param>
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
