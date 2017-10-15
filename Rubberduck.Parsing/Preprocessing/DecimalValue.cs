﻿using System;
using System.Globalization;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class DecimalValue : IValue
    {
        private readonly decimal _value;

        public DecimalValue(decimal value)
        {
            _value = value;
        }

        public ValueType ValueType
        {
            get
            {
                return ValueType.Decimal;
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
                return Convert.ToByte(_value);
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
                return _value;
            }
        }

        public string AsString
        {
            get
            {
                return _value.ToString(CultureInfo.InvariantCulture);
            }
        }

        public override string ToString()
        {
            return _value.ToString(CultureInfo.InvariantCulture);
        }
    }
}
