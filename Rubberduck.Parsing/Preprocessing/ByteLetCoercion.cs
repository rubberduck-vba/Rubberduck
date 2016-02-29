using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class ByteLetCoercion : ILetCoercion
    {
        public bool ToBool(object value)
        {
            return (byte)value != 0;
        }

        public byte ToByte(object value)
        {
            return (byte)value;
        }

        public DateTime ToDate(object value)
        {
            return DateTime.FromOADate(Convert.ToDouble(value));
        }

        public decimal ToDecimal(object value)
        {
            return Convert.ToDecimal((byte)value);
        }

        public string ToString(object value)
        {
            return value.ToString();
        }
    }
}
