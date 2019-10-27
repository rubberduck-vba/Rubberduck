using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using NLog;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeComWrapper<T> : ISafeComWrapper<T>
        where T : class
    {
        protected static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private IComSafe _comSafe;
        private bool _rewrapping;

        protected SafeComWrapper(T target, bool rewrapping = false)
        {
            Target = target;
            _rewrapping = rewrapping;

            if (!rewrapping && target != null)
            {
                _comSafe = ComSafeManager.GetCurrentComSafe();
                _comSafe.Add(this);
            }
        }

        private int? _rcwReferenceCount;
        private void Release(bool final = false)
        {
            if (HasBeenReleased)
            {
                _logger.Warn($"Tried to release already released COM wrapper of type {this.GetType()}.");
                return;
            }
            if (IsWrappingNullReference)
            {
                _rcwReferenceCount = 0;
                _logger.Warn($"Tried to release a COM wrapper of type {this.GetType()} wrapping a null reference.");
                return;
            }

            if (!Marshal.IsComObject(Target))
            {
                _rcwReferenceCount = 0;
                _logger.Warn($"Tried to release a COM wrapper of type {this.GetType()} whose target is not a COM object.");
                return;
            }

            try
            {
                if (final)
                {
                    _rcwReferenceCount = Marshal.FinalReleaseComObject(Target);
                    if (HasBeenReleased)
                    {
                        _logger.Trace($"Final released COM wrapper of type {this.GetType()}.");
                    }
                    else
                    {
                        _logger.Warn($"Final released COM wrapper of type {this.GetType()} did not release the wrapper: remaining reference count is {_rcwReferenceCount}.");
                    }
                }
                else
                {
                    _rcwReferenceCount = Marshal.ReleaseComObject(Target);
                    if (_rcwReferenceCount < 0)
                    {
                        _logger.Warn($"Released COM wrapper of type {this.GetType()} whose underlying RCW has already been released from outside the SafeComWrapper. New reference count is {_rcwReferenceCount}.");
                    }
                    else
                    {
                        LogComRelease();
                    }
                }
            }
            catch(COMException exception)
            {
                var logMessage = $"Failed to release COM wrapper of type {this.GetType()}.";
                if (_rcwReferenceCount.HasValue)
                {
                    logMessage = logMessage + $"The previous reference count has been {_rcwReferenceCount}.";
                }
                else
                {
                    logMessage = logMessage + "There has not yet been an attempt to release the COM wrapper.";
                }

                _logger.Warn(exception, logMessage);
            }
        }

        private bool HasBeenReleased => _rcwReferenceCount <= 0;

        public bool IsWrappingNullReference => Target == null;
        object INullObjectWrapper.Target => Target;
        public T Target { get; }

        /// <summary>
        /// <c>true</c> when wrapping a <c>null</c> reference and <see cref="other"/> is either <c>null</c> or wrapping a <c>null</c> reference.
        /// </summary>
        protected bool IsEqualIfNull(ISafeComWrapper<T> other)
        {
            return (other == null || other.IsWrappingNullReference) && IsWrappingNullReference;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as ISafeComWrapper<T>);
        }

        public abstract bool Equals(ISafeComWrapper<T> other);
        public abstract override int GetHashCode();

        public static bool operator ==(SafeComWrapper<T> a, SafeComWrapper<T> b)
        {
            return ReferenceEquals(a, null) ? ReferenceEquals(b, null) : a.Equals(b);
        }

        public static bool operator !=(SafeComWrapper<T> a, SafeComWrapper<T> b)
        {
            return !(a == b);
        }


        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private readonly object _disposalLockObject = new object();
        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            lock (_disposalLockObject)
            {
                if (_isDisposed)
                {
                    return;
                }
                _isDisposed = true;
            }

            if (disposing)
            {
                _comSafe?.TryRemove(this);

                if (!_rewrapping && !IsWrappingNullReference && !HasBeenReleased)
                {
                    Release();
                }

                _comSafe = null;
            }
        }

        [Conditional("LOG_COM_RELEASE")]
        private void LogComRelease()
        {
            _logger.Trace($"Released COM wrapper of type {this.GetType()} with remaining reference count {_rcwReferenceCount}.");
        }
    }
}
