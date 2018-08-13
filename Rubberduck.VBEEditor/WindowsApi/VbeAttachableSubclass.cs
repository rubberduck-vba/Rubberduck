using System;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    public interface ISubclassAttachable : ISafeComWrapper { }

    internal abstract class VbeAttachableSubclass<T> : FocusSource where T : ISubclassAttachable
    {
        protected VbeAttachableSubclass(IntPtr hWnd) : base(hWnd) { }

        private T _target;

        /// <summary>
        /// The VbeObject associated with the message pump (if it has successfully been found).
        /// WARNING: Internal callers should NOT call *anything* on this object. Remember, you're in its message pump here.
        /// External callers should NOT call .Dispose() on this object. That's the VbeAttachableSubclass's responsibility.
        /// </summary>
        public T VbeObject
        {
            get => _target;
            set
            {
                if (HasValidVbeObject)
                {
                    _target.Dispose();
                }

                _target = value;
            }
        }

        /// <summary>
        /// Returns true if the Subclass is:
        /// 1.) Holding a VbeObject reference
        /// 2.) The held reference is pointed to a valid object (i.e. it has not been recycled). 
        /// </summary>
        public bool HasValidVbeObject
        {
            get
            {
                if (_target == null)
                {
                    return false;
                }

                try
                {
                    var unmanaged = Marshal.GetIUnknownForObject(_target.Target);
                    if (unmanaged != IntPtr.Zero)
                    {
                        try
                        {
                            Marshal.Release(unmanaged);
                        }
                        catch (COMException)
                        {
                            // this should be "safe" to hold the reference - it's not like it should go anywhere now... :-/
                            SubclassLogger.Warn($"{ GetType().Name } failed to decrement reference count of a held { typeof(T).Name } reference.");
                        }
                        return true;
                    }
                }
                catch (COMException)
                {
                    _target = default;
                }
                return false;
            }
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                if (HasValidVbeObject)
                {
                    _target.Dispose();
                    _target = default;
                }
            }

            base.Dispose(disposing);
            _disposed = true;
        }
    }
}
