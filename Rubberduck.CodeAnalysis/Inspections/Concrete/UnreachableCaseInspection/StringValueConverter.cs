using Rubberduck.Parsing.Grammar;
using System;
using System.Globalization;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public delegate bool StringToValueConversion<T>(string value, out T result, string typeName = null);

    public class StringValueConverter
    {
        public static bool TryConvertString(string valueText, out long value, string typeName = null)
        {
            value = default;
            typeName = typeName ?? Tokens.Long;

            if (valueText.Equals(Tokens.True) || valueText.Equals(Tokens.False))
            {
                value = valueText.Equals(Tokens.True) ? -1 : 0;
                return true;
            }

            if ((typeName == Tokens.Byte
                    || typeName == Tokens.Integer
                    || typeName == Tokens.Long
                    || typeName == Tokens.LongLong)
                && long.TryParse(valueText, out var integralValue))
            {
                value = integralValue;
                return true;
            }

            if (typeName == Tokens.Currency && decimal.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalValue))
            {
                value = Convert.ToInt64(decimalValue);
                return true;
            }

            if (double.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var rationalValue))
            {
                value = Convert.ToInt64(rationalValue);
                return true;
            }

            return false;
        }

        public static bool TryConvertString(string valueText, out double value, string typeName = null)
        {
            value = default;
            if (valueText.Equals(Tokens.True) || valueText.Equals(Tokens.False))
            {
                value = valueText.Equals(Tokens.True) ? -1 : 0;
                return true;
            }
            if (double.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var rational))
            {
                value = rational;
                return true;
            }
            return false;
        }

        public static bool TryConvertString(string valueText, out decimal value, string typeName = null)
        {
            value = default;
            if (valueText.Equals(Tokens.True) || valueText.Equals(Tokens.False))
            {
                value = valueText.Equals(Tokens.True) ? -1 : 0;
                return true;
            }

            if (decimal.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var decimalValue))
            {
                value = decimalValue;
                return true;
            }

            if (double.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var rationalValue))
            {
                value = Convert.ToDecimal(rationalValue);
                return true;
            }

            return false;
        }

        public static bool TryConvertString(string valueText, out bool value, string typeName = null)
        {
            value = default;
            if (valueText.Equals(Tokens.True) || valueText.Equals(Tokens.False))
            {
                value = valueText.Equals(Tokens.True);
                return true;
            }
            if (double.TryParse(valueText, NumberStyles.Any, CultureInfo.InvariantCulture, out var doubleValue))
            {
                value = Math.Abs(doubleValue) >= double.Epsilon;
                return true;
            }
            return false;
        }

        public static bool TryConvertString(string valueText, out string value, string typeName = null)
        {
            value = valueText;
            return true;
        }
    }
}
