using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public delegate bool TryConvertParseTreeValue<T>(IParseTreeValue value, out T result);

    internal class ParseTreeValueConverter
    {
        internal static long ConvertLong(IParseTreeValue value)
        {
            return ConvertLong(value.ValueText);
        }

        //public static bool TryConvertValue(IParseTreeValue value, out long result)
        //{
        //    return TryConvertValue(value.ValueText, out result);
        //}

        private static long ConvertLong(string value)
        {
            if (TryConvertValue(value, out long result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value}) to {result.GetType()}");
        }

        internal static double ConvertDouble(IParseTreeValue value)
        {
            return ConvertDouble(value.ValueText);
        }

        public static bool TryConvertValue(IParseTreeValue value, out double result)
        {
            return TryConvertValue(value.ValueText, out result);
        }

        private static double ConvertDouble(string value)
        {
            if (TryConvertValue(value, out double result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value}) to {result.GetType()}");
        }

        internal static decimal ConvertDecimal(IParseTreeValue value)
        {
            return ConvertDecimal(value.ValueText);
        }

        private static decimal ConvertDecimal(string value)
        {
            if (TryConvertValue(value, out decimal result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value}) to {result.GetType()}");
        }

        internal static bool ConvertBoolean(IParseTreeValue value)
        {
            return ConvertBoolean(value.ValueText);
        }

        private static bool ConvertBoolean(string value)
        {
            if (TryConvertValue(value, out bool result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value}) to {result.GetType()}");
        }

        internal static string ConvertString(IParseTreeValue value)
        {
            return value.ValueText;
        }

        internal static bool TryConvertValue(string inspVal, out long value)
        {
            value = default;
            if (inspVal.Equals(Tokens.True) || inspVal.Equals(Tokens.False))
            {
                value = inspVal.Equals(Tokens.True) ? -1 : 0;
                return true;
            }
            if (double.TryParse(inspVal, out double rational))
            {
                value = Convert.ToInt64(rational);
                return true;
            }
            return false;
        }

        internal static bool TryConvertValue(string inspVal, out double value)
        {
            value = default;
            if (inspVal.Equals(Tokens.True) || inspVal.Equals(Tokens.False))
            {
                value = inspVal.Equals(Tokens.True) ? -1 : 0;
                return true;
            }
            if (double.TryParse(inspVal, out double rational))
            {
                value = rational;
                return true;
            }
            return false;
        }

        internal static bool TryConvertValue(string inspVal, out decimal value)
        {
            value = default;
            if (inspVal.Equals(Tokens.True) || inspVal.Equals(Tokens.False))
            {
                value = inspVal.Equals(Tokens.True) ? -1 : 0;
                return true;
            }
            if (decimal.TryParse(inspVal, out decimal rational))
            {
                value = rational;
                return true;
            }
            return false;
        }

        internal static bool TryConvertValue(string inspVal, out bool value)
        {
            value = default;
            if (inspVal.Equals(Tokens.True) || inspVal.Equals(Tokens.False))
            {
                value = inspVal.Equals(Tokens.True);
                return true;
            }
            if (double.TryParse(inspVal, out double dVal))
            {
                value = Math.Abs(dVal) > double.Epsilon;
                return true;
            }
            return false;
        }
    }
}
