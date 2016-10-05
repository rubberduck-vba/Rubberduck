using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.SafeComWrappers
{
    public abstract class SafeComWrapper<T> : ISafeComWrapper, IEquatable<SafeComWrapper<T>> 
        where T : class 
    {
        protected SafeComWrapper(T comObject)
        {
            _comObject = comObject;
        }

        public abstract void Release();

        private readonly T _comObject;
        public T ComObject { get { return _comObject; } }
        public bool IsWrappingNullReference { get { return _comObject == null; } }

        protected TResult InvokeResult<TResult>(Func<TResult> member)
        {
            try
            {
                return member.Invoke();
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
            catch (NullReferenceException exception)
            {
                throw new WrapperMethodException(exception);                
            }
        }

        protected void Invoke(Action member)
        {
            try
            {
                member.Invoke();
            }
            catch (COMException exception)
            {
                throw new WrapperMethodException(exception);
            }
            catch (NullReferenceException exception)
            {
                throw new WrapperMethodException(exception);
            }
        }

        /// <summary>
        /// <c>true</c> when wrapping a <c>null</c> reference and <see cref="other"/> is either <c>null</c> or wrapping a <c>null</c> reference.
        /// </summary>
        protected bool IsEqualIfNull(SafeComWrapper<T> other)
        {
            return (other == null || other.IsWrappingNullReference) && IsWrappingNullReference;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as SafeComWrapper<T>);
        }

        public abstract bool Equals(SafeComWrapper<T> other);
        public abstract override int GetHashCode();

        public static bool operator ==(SafeComWrapper<T> a, SafeComWrapper<T> b)
        {
            if (ReferenceEquals(a, null) && ReferenceEquals(b, null))
            {
                return true;
            }
            return !ReferenceEquals(a, null) && a.Equals(b);
        }

        public static bool operator !=(SafeComWrapper<T> a, SafeComWrapper<T> b)
        {
            return !(a == b);
        }

        [SuppressMessage("ReSharper", "RedundantCast")]
        [SuppressMessage("ReSharper", "ForCanBeConvertedToForeach")]
        [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
        protected int ComputeHashCode(params object[] values) // incurs boxing penalty for value types
        {
            unchecked
            {
                const int initial = (int)2166136261;
                const int multiplier = (int)16777619;
                var hash = initial;
                for (var i = 0; i < values.Length; i++)
                {
                    hash = hash * multiplier + values[i].GetHashCode();
                }
                return hash;
            }
        }
   }
}