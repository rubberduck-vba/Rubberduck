using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionValueConvert
    {
        string ConvertToType(string value, string typeName);
    }

    internal class UCIValueConverter
    {
        public static long ConvertLong(IUnreachableCaseInspectionValue value)
        {
            if (TryConvert(value, out long result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value.ValueText}) to {result.GetType()}");
        }

        public static double ConvertDouble(IUnreachableCaseInspectionValue value)
        {
            if (TryConvert(value, out double result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value.ValueText}) to {result.GetType()}");
        }

        public static decimal ConvertDecimal(IUnreachableCaseInspectionValue value)
        {
            if (TryConvert(value, out decimal result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value.ValueText}) to {result.GetType()}");
        }

        public static bool ConvertBoolean(IUnreachableCaseInspectionValue value)
        {
            if (TryConvert(value, out bool result))
            {
                return result;
            }
            throw new ArgumentException($"Unable to convert parameter (value = {value.ValueText}) to {result.GetType()}");
        }

        public static string ConvertString(IUnreachableCaseInspectionValue value)
        {
            return value.ValueText;
        }

        internal static double? GetOperandAsDouble(IUnreachableCaseInspectionValue value)
        {
            double? result;
            var conformed = new UnreachableCaseInspectionValueConformed(value, value.TypeName);
            if (!TryConvert(conformed, out double conformedValue))
            {
                result = null;
            }
            else
            {
                result = conformedValue;
            }
            return result;
        }

        internal static IUnreachableCaseInspectionValue ConvertToType<T>(T tValue, string targetType)
        {
            try { tValue.ToString(); }
            catch (NullReferenceException)
            {
                return new UnreachableCaseInspectionValue(double.NaN.ToString(), targetType);
            }

            if (UnreachableCaseInspectionValue.IntegerTypes.Contains(targetType))
            {
                if (TryConvertValue(tValue, out long result))
                {
                    return new UnreachableCaseInspectionValue(result);
                }
            }

            if (UnreachableCaseInspectionValue.RationalTypes.Contains(targetType))
            {
                if (targetType.Equals(Tokens.Currency))
                {
                    if (TryConvertValue(tValue, out decimal cResult))
                    {
                        return new UnreachableCaseInspectionValue(cResult);
                    }
                }

                if (TryConvertValue(tValue, out double dResult))
                {
                    return new UnreachableCaseInspectionValue(dResult);
                }
            }

            if (targetType.Equals(Tokens.Boolean))
            {
                if (tValue.ToString().Equals(Tokens.True) || tValue.ToString().Equals(Tokens.False))
                {
                    var value = tValue.ToString().Equals(Tokens.True) ? -1 : 0;
                    return new UnreachableCaseInspectionValue(value != 0); ;
                }
                if (TryConvertValue(tValue, out long result))
                {
                    return new UnreachableCaseInspectionValue(result != 0);
                }
            }

            if (targetType.Equals(Tokens.String))
            {
                return new UnreachableCaseInspectionValue($"\"{tValue.ToString()}\"", Tokens.String);
            }
            return new UnreachableCaseInspectionValue(double.NaN.ToString(), targetType);
        }

        static private bool TryConvert(IUnreachableCaseInspectionValue valToConvert, out long result)
        {
            result = default;
            if (!TryConvertValue(valToConvert.ValueText, out result))
            {
                if (valToConvert.TypeName.Equals(Tokens.Boolean))
                {
                    result = valToConvert.ValueText.Equals(Tokens.True) ? -1 : 0;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        static private bool TryConvert(IUnreachableCaseInspectionValue valToConvert, out double result)
        {
            result = default;
            if (!TryConvertValue(valToConvert.ValueText, out result))
            {
                if (valToConvert.TypeName.Equals(Tokens.Boolean))
                {
                    result = valToConvert.ValueText.Equals(Tokens.True) ? -1.0 : 0.0;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        static private bool TryConvert(IUnreachableCaseInspectionValue valToConvert, out decimal result)
        {
            result = default;
            if (!TryConvertValue(valToConvert.ValueText, out result))
            {
                if (valToConvert.TypeName.Equals(Tokens.Boolean))
                {
                    result = valToConvert.ValueText.Equals(Tokens.True) ? -1.0M : 0.0M;
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        static private bool TryConvert(IUnreachableCaseInspectionValue valToConvert, out bool result)
        {
            result = default;
            if (!TryConvertValue(valToConvert.ValueText, out result))
            {
                if (valToConvert.TypeName.Equals(Tokens.Boolean))
                {
                    result = valToConvert.ValueText.Equals(Tokens.True);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        private static bool TryConvertValue<T>(T inspVal, out long value)
        {
            value = default;
            if (typeof(T) == typeof(bool))
            {
                value = inspVal.ToString() == bool.TrueString ? -1 : 0;
                return true;
            }

            if (double.TryParse(inspVal.ToString(), out double rational))
            {
                value = Convert.ToInt64(rational);
                return true;
            }
            return false;
        }

        private static bool TryConvertValue<T>(T inspVal, out double value)
        {
            value = default;
            if (typeof(T) == typeof(bool))
            {
                value = inspVal.ToString() == bool.TrueString ? -1.0 : 0.0;
                return true;
            }

            if (double.TryParse(inspVal.ToString(), out double rational))
            {
                value = rational;
                return true;
            }
            return false;
        }

        private static bool TryConvertValue<T>(T inspVal, out decimal value)
        {
            value = default;
            if (decimal.TryParse(inspVal.ToString(), out decimal rational))
            {
                value = rational;
                return true;
            }
            return false;
        }

        private static bool TryConvertValue<T>(T inspVal, out bool value)
        {
            value = default;
            if (bool.TryParse(inspVal.ToString(), out bool booleanVal))
            {
                value = booleanVal;
                return true;
            }
            if(long.TryParse(inspVal.ToString(), out long lValue))
            {
                value = lValue != 0 ? true : false;
                return true;
            }
            return false;
        }
    }
}
