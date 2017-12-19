using System.Runtime.InteropServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using NLog;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeComWrapper<T> : ISafeComWrapper<T>
        where T : class
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();     

        protected SafeComWrapper(T target)
        {
            Target = target;
        }

        private int? _rcwReferenceCount;
        public virtual void Release(bool final = false)
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
                    _logger.Trace($"Released COM wrapper of type {this.GetType()} with remaining reference count {_rcwReferenceCount}.");
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

        public bool HasBeenReleased => _rcwReferenceCount == 0;

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
   }
}
