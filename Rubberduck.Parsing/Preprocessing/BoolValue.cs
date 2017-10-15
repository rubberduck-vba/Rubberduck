using System;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class BoolValue : IValue
    {
        private readonly bool _value;

        public BoolValue(bool value)
        {
            _value = value;
        }

        public ValueType ValueType
        {
            get
            {
                return ValueType.Bool;
            }
        }

        public bool AsBool
        {
            get
            {
                return _value;
            }
        }

        public byte AsByte
        {
            get
            {
                if (_value)
                {
                    return 255;
                }
                return 0;
            }
        }

        public DateTime AsDate
        {
            get
            {
                return new DecimalValue(AsDecimal).AsDate;
            }
        }

        public decimal AsDecimal
        {
            get
            {
                if (_value)
                {
                    return -1;
                }
                else
                {
                    return 0;
                }
            }
        }

        public string AsString
        {
            get
            {
                if (_value)
                {
                    return "True";
                }
                else
                {
                    return "False";
                }
            }
        }

        public override string ToString()
        {
            return _value.ToString();
        }
    }
}
