using Antlr4.Runtime;
using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class ByteValue : IValue
    {
        private readonly byte _value;

        public ByteValue(byte value)
        {
            _value = value;
        }

        public ValueType ValueType
        {
            get
            {
                return ValueType.Byte;
            }
        }

        public bool AsBool
        {
            get
            {
                return _value != 0;
            }
        }

        public byte AsByte
        {
            get
            {
                return _value;
            }
        }

        public DateTime AsDate
        {
            get
            {
                return DateTime.FromOADate(Convert.ToDouble(_value));
            }
        }

        public decimal AsDecimal
        {
            get
            {
                return Convert.ToDecimal(_value);
            }
        }

        public string AsString
        {
            get
            {
                return _value.ToString();
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
            return _value.ToString();
        }
    }
}
