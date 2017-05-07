using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class TokensValue : IValue
    {
        private readonly IEnumerable<IToken> _value;

        public TokensValue(IEnumerable<IToken> value)
        {
            _value = value;
        }

        public ValueType ValueType
        {
            get
            {
                return ValueType.Tokens;
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
                var value = AsString;
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
                if (byte.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    return value;
                }
                return byte.Parse(AsString, NumberStyles.Float);
            }
        }

        public DateTime AsDate
        {
            get
            {
                DateTime value;
                if (DateTime.TryParse(AsString, out value))
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
                if (decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
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
                var builder = new StringBuilder();
                foreach (var token in _value)
                {
                    if (token.Channel == TokenConstants.DefaultChannel)
                    {
                        builder.Append(token.Text);
                    }
                }
                return builder.ToString();
            }
        }

        public IEnumerable<IToken> AsTokens
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
