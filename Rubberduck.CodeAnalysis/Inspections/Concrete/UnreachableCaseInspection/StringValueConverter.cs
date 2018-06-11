using Rubberduck.Parsing.Grammar;
using System;
using System.Globalization;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    //The 'typeName' parameter is only used for Integral types, but since this delegate is assigned 
    //to ExpressionFilters (a generic class) during construction, the signatures for the
    //non-integral types have to match
    public delegate bool StringToValueConversion<T>(string value, string typeName, out T result);

    public class StringValueConverter
    {
        public static bool TryConvertString(string valueText, string typeName, out long value)
        {
            value = default;
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

        public static bool TryConvertString(string valueText, string typeName, out double value)
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

        public static bool TryConvertString(string valueText, string typeName, out decimal value)
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

        public static bool TryConvertString(string valueText, string typeName, out bool value)
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

        public static bool TryConvertString(string valueText, string typeName, out string value)
        {
            value = valueText;
            return true;
        }
    }
}
