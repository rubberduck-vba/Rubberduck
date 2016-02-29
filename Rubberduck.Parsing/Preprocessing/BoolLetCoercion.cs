using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class BoolLetCoercion : ILetCoercion
    {
        public bool ToBool(object value)
        {
            return (bool)value;
        }

        public byte ToByte(object value)
        {
            if ((bool)value)
            {
                return 255;
            }
            return 0;
        }

        public DateTime ToDate(object value)
        {
            return new DecimalLetCoercion().ToDate(ToDecimal(value));
        }

        public decimal ToDecimal(object value)
        {
            if ((bool)value)
            {
                return -1;
            }
            else
            {
                return 0;
            }
        }

        public string ToString(object value)
        {
            if ((bool)value)
            {
                return "True";
            }
            else
            {
                return "False";
            }
        }
    }
}
