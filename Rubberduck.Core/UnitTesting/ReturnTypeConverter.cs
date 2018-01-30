using System;

namespace Rubberduck.UnitTesting
{
    internal class ValueTypeConverter<T> where T : struct
    {
        public bool IsValid => _value.HasValue;

        private T? _value;
#pragma warning disable 649
        private T _default;     //not assigned because we *want* the default T.
#pragma warning restore 649
        public object Value
        {
            get { return _value ?? _default; }
            set
            {
                _value = value.GetType().IsPrimitive ? (T) Convert.ChangeType(value, typeof(T)) : null as T?;
            }
        }
    }
}
