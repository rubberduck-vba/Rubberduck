using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class DecimalLetCoercion : ILetCoercion
    {
        public bool ToBool(object value)
        {
            return (decimal)value != 0;
        }

        public byte ToByte(object value)
        {
            return Convert.ToByte(value);
        }

        public DateTime ToDate(object value)
        {
            return DateTime.FromOADate(Convert.ToDouble(value));
        }

        public decimal ToDecimal(object value)
        {
            return (decimal)value;
        }

        public string ToString(object value)
        {
            return value.ToString();
        }
    }
}
