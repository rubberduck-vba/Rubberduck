using System;
using System.Globalization;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class StringValue : IValue
    {
        private readonly string _value;

        public StringValue(string value)
        {
            _value = value;
        }

        public ValueType ValueType
        {
            get
            {
                return ValueType.String;
            }
        }

        public bool AsBool
        {
            get
            {
                if (_value == null)
                {
                    return false;
                }
                var str = _value;
                if (string.CompareOrdinal(str.ToLower(), "true") == 0
                    || string.CompareOrdinal(str, "#TRUE#") == 0)
                {
                    return true;
                }
                else if (string.CompareOrdinal(str.ToLower(), "false") == 0
                    || string.CompareOrdinal(str, "#FALSE#") == 0)
                {
                    return false;
                }
                else
                {
                    decimal number = AsDecimal;
                    return new DecimalValue(number).AsBool;
                }
            }
        }

        public byte AsByte
        {
            get
            {
                return byte.Parse(_value, NumberStyles.Float);
            }
        }

        public DateTime AsDate
        {
            get
            {
                DateTime date;
                if (DateTime.TryParse(_value, out date))
                {
                    return date;
                }
                decimal number = AsDecimal;
                return new DecimalValue(number).AsDate;
            }
        }

        public decimal AsDecimal
        {
            get
            {
                return decimal.Parse(_value, NumberStyles.Float);
            }
        }

        public string AsString
        {
            get
            {
                return _value;
            }
        }

        public override string ToString()
        {
            return _value.ToString();
        }
    }
}
