using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Data;
using System.Globalization;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public delegate bool TokenToValue<T>(string value, out T result, string typeName = null);

    public class TokenParser
    {
        public static bool TryParse(string valueText, out long value, string typeName = null)
        {
            value = default;
            valueText = StripDoubleQuotes(valueText);
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

        public static bool TryParse(string valueText, out double value, string typeName = null)
        {
            value = default;
            valueText = StripDoubleQuotes(valueText);
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

        public static bool TryParse(string valueText, out decimal value, string typeName = null)
        {
            value = default;
            valueText = StripDoubleQuotes(valueText);
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

            return false;
        }

        public static bool TryParse(string valueText, out bool value, string typeName = null)
        {
            value = default;
            valueText = StripDoubleQuotes(valueText);
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

        public static bool TryParse(string valueString, out ComparableDateValue value, string typeName = null)
        {
            value = default;
            if (!(valueString.StartsWith("#") && valueString.EndsWith("#")))
            {
                return false;
            }

            try
            {
                var literal = new DateLiteralExpression(new ConstantExpression(new StringValue(valueString)));
                value = new ComparableDateValue((DateValue)literal.Evaluate());
                return true;
            }
            catch (SyntaxErrorException)
            {
                return false;
            }
            catch (Exception)
            {
                //even though a SyntaxErrorException is thrown, this catch block
                //seems to be needed(?)
                return false;
            }
        }

        public static bool TryParse(string valueText, out string value, string typeName = null)
        {
            value = valueText;
            return true;
        }

        private static string StripDoubleQuotes(string input)
        {
            if (input.StartsWith("\""))
            {
                input = input.Substring(1);
            }
            if (input.EndsWith("\""))
            {
                input = input.Substring(0,input.Length - 1);
            }
            return input;
        }
    }
}
