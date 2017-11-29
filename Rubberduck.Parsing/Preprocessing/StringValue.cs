using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class StringValue : IValue
    {
        public StringValue(string value)
        {
            AsString = value;
        }

        public ValueType ValueType => ValueType.String;

        public bool AsBool
        {
            get
            {
                if (AsString == null)
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

        public byte AsByte => byte.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) 
            ? value 
            : byte.Parse(AsString, NumberStyles.Float);

        public DateTime AsDate
        {
            get
            {
                if (DateTime.TryParse(AsString, out var value))
                {
                    return value;
                }
                var number = AsDecimal;
                return new DecimalValue(number).AsDate;
            }
        }

        public decimal AsDecimal
        {
            get
            {
                if (decimal.TryParse(AsString, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                {
                    return value;
                }
                Debug.Assert(false); // this line was never hit in any unit test covering it.
                return 0;
            }
        }

        public string AsString { get; }

        public IEnumerable<IToken> AsTokens => new List<IToken>();

        public override string ToString()
        {
            return AsString;
        }
    }
}
