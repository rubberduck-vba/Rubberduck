using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeComWrapper<T> : ISafeComWrapper<T>
        where T : class
    {
        protected SafeComWrapper(T target)
        {
            Target = target;
        }

        //private bool _isReleased;
        //public virtual void Release(bool final = false)
        //{
        //    if (IsWrappingNullReference || _isReleased || !Marshal.IsComObject(Target))
        //    {
        //        _isReleased = true;
        //        return;
        //    }

        //    try
        //    {
        //        if (final)
        //        {
        //            Marshal.FinalReleaseComObject(Target);
        //        }
        //        else
        //        {
        //            Marshal.ReleaseComObject(Target);
        //        }
        //    }
        //    finally
        //    {
        //        _isReleased = true;
        //    }
        //}

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
