using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.PreProcessing
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
                var value = _value;
                if (string.CompareOrdinal(value.ToLower(), "true") == 0 || string.CompareOrdinal(value, "#TRUE#") == 0)
                {
                    return true;
                }
                
                if (string.CompareOrdinal(value.ToLower(), "false") == 0 || string.CompareOrdinal(value, "#FALSE#") == 0)
                {
                    return false;
                }
                
                return new DecimalValue(AsDecimal).ToString() != "0"; // any non-zero value evaluates to TRUE in VBA
            }
        }

        public byte AsByte
        {
            get
            {
                byte value;
                if (byte.TryParse(_value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }
                return byte.Parse(_value, NumberStyles.Float);
            }
        }

        public DateTime AsDate
        {
            get
            {
                DateTime value;
                if (DateTime.TryParse(_value, out value))
                {
                    return value;
                }
                decimal number = AsDecimal;
                return new DecimalValue(number).AsDate;
            }
        }

        public decimal AsDecimal
        {
            get
            {
                decimal value;
                if (decimal.TryParse(_value, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }
                Debug.Assert(false); // this line was never hit in any unit test covering it.
                return 0;
            }
        }

        public string AsString
        {
            get
            {
                return _value;
            }
        }

        public IEnumerable<IToken> AsTokens
        {
            get
            {
                return new List<IToken>();
            }
        }

        public override string ToString()
        {
            return _value;
        }
    }
}
