using System;
using System.Globalization;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class StringLetCoercion : ILetCoercion
    {
        public bool ToBool(object value)
        {
            if (value == null)
            {
                return false;
            }
            var str = (string)value;
            if (str.ToLower() == "true" || str == "#TRUE#")
            {
                return true;
            }
            else if (str.ToLower() == "false" || str == "#FALSE#")
            {
                return false;
            }
            else
            {
                decimal number = ToDecimal(value);
                return new DecimalLetCoercion().ToBool(number);
            }
        }

        public byte ToByte(object value)
        {
            return byte.Parse((string)value, NumberStyles.Float);
        }

        public DateTime ToDate(object value)
        {
            DateTime date;
            if (DateTime.TryParse((string)value, out date))
            {
                return date;
            }
            decimal number = ToDecimal(value);
            return new DecimalLetCoercion().ToDate(number);
        }

        public decimal ToDecimal(object value)
        {
            return decimal.Parse((string)value, NumberStyles.Float);
        }

        public string ToString(object value)
        {
            return (string)value;
        }
    }
}
