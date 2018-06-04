using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//TODO: This class may need to go away - compare and update per IParseTreeValueExtensions class
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public delegate bool StringToValueConversion<T>(string value, out T result);

    public class StringValueConverter
    {
        public static bool TryConvertString(string valueAsText, out long value)
        {
            value = default;
            if (valueAsText.Equals(Tokens.True) || valueAsText.Equals(Tokens.False))
            {
                value = valueAsText.Equals(Tokens.True) ? -1 : 0;
                return true;
            }
            if (double.TryParse(valueAsText, out double rational))
            {
                //protect against double.NaN
                try
                {
                    value = Convert.ToInt64(rational);
                    return true;
                }
                catch (OverflowException) { }
            }
            return false;
        }

        public static bool TryConvertString(string inspVal, out double value)
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

        public static bool TryConvertString(string inspVal, out decimal value)
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

        public static bool TryConvertString(string inspVal, out bool value)
        {
            value = default;
            if (inspVal.Equals(Tokens.True) || inspVal.Equals(Tokens.False))
            {
                value = inspVal.Equals(Tokens.True);
                return true;
            }
            if (double.TryParse(inspVal, out double dVal))
            {
                value = Math.Abs(dVal) >= double.Epsilon;
                return true;
            }
            return false;
        }

        public static bool TryConvertString(string inspVal, out string value)
        {
            value = inspVal;
            return true;
        }
    }
}
