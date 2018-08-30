using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class TokenTypeResolver
    {
        public static string DeriveTypeName(string inputString, out bool derivedFromTypeHint)
        {
            if (TryDeriveTypeName(inputString, out (string TypeName, string Value) result, out derivedFromTypeHint))
            {
                return result.TypeName;
            }
            return string.Empty;
        }

        public static bool TryDeriveTypeName(string inputString, out (string typeName, string valueText) result, out bool derivedFromTypeHint)
        {
            string typeHintType = null;
            derivedFromTypeHint = false;
            result.typeName = string.Empty;
            result.valueText = string.Empty;

            if (inputString is null || inputString.Length == 0)
            {
                return false;
            }

            if (TryParseAsDateLiteral(inputString, out ComparableDateValue cdv))
            {
                result = (Tokens.Date, cdv.AsString);
                return true;
            }

            if (SymbolList.TypeHintToTypeName.TryGetValue(inputString.Last().ToString(), out string hintResult))
            {
                derivedFromTypeHint = true;
                typeHintType = hintResult;
                inputString = inputString.Remove(inputString.Length - 1);
                result = (typeHintType, inputString);
            }

            if (typeHintType == Tokens.String)
            {
                return true;
            }


            if (IsStringConstant(inputString))
            {
                result = (Tokens.String, inputString);
                return true;
            }

            if (typeHintType == null || typeHintType == Tokens.Single)
            {
                if (inputString.Contains(".") || inputString.Count(ch => ch.Equals('E')) == 1)
                {
                    if (double.TryParse(inputString, NumberStyles.Float, CultureInfo.InvariantCulture, out double dVal))
                    {
                        result = (typeHintType ?? Tokens.Double, dVal.ToString());
                        return true;
                    }
                }
            }

            if (inputString.Equals(Tokens.True) || inputString.Equals(Tokens.False))
            {
                result = (Tokens.Boolean, inputString);
                return true;
            }

            if (typeHintType == null || typeHintType == Tokens.Integer)
            {
                if (short.TryParse(inputString, out short shVal)
                    || TryParseAsHexLiteral(inputString, out shVal)
                    || TryParseAsOctalLiteral(inputString, out shVal))
                {
                    result = (typeHintType ?? Tokens.Integer, shVal.ToString());
                    return true;
                }
            }

            if (typeHintType == null || typeHintType == Tokens.Long)
            {
                if (int.TryParse(inputString, out int intVal)
                    || TryParseAsHexLiteral(inputString, out intVal)
                    || TryParseAsOctalLiteral(inputString, out intVal))
                {
                    result = (typeHintType ?? Tokens.Long, intVal.ToString());
                    return true;
                }
            }
            if (typeHintType == null || typeHintType == Tokens.LongLong)
            {
                if (TryParseAsHexLiteral(inputString, out long outputHex))
                {
                    result = (typeHintType ?? Tokens.Double, outputHex.ToString());
                    return true;
                }

                if (TryParseAsOctalLiteral(inputString, out long outputOctal))
                {
                    result = (typeHintType ?? Tokens.Double, outputOctal.ToString());
                    return true;
                }
            }

            if (long.TryParse(inputString, out long lngVal))
            {
                result = (typeHintType ?? Tokens.Double, lngVal.ToString());
                return true;
            }

            return derivedFromTypeHint;
        }

        public static bool TryConformTokenToType(string valueToken, string conformTypeName, out string result)
        {
            result = string.Empty;

            if (valueToken.Equals(double.NaN.ToString(CultureInfo.InvariantCulture)) &&
                !conformTypeName.Equals(Tokens.String))
            {
                return false;
            }

            else if (conformTypeName.Equals(Tokens.LongLong) || conformTypeName.Equals(Tokens.Long) ||
                conformTypeName.Equals(Tokens.Integer) || conformTypeName.Equals(Tokens.Byte))
            {
                //If a coercion fails, it may be an overflow condition for the conformTypeName
                var isArabicNumber = LetCoercer.TryCoerce(valueToken, out long checkValue);

                if (conformTypeName.Equals(Tokens.Byte))
                {
                    if (LetCoercer.TryCoerce(valueToken, out byte newVal))
                    {
                        result = newVal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }
                }
                else if (conformTypeName.Equals(Tokens.Integer))
                {
                    if (LetCoercer.TryCoerce(valueToken, out short newVal))
                    {
                        result = newVal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    if (TryParseAsHexLiteral(valueToken, out short outputHex))
                    {
                        result = outputHex.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out short outputOctal))
                    {
                        result = outputOctal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }
                }
                else if (conformTypeName.Equals(Tokens.Long))
                {
                    if (LetCoercer.TryCoerce(valueToken, out int newVal))
                    {
                        result = newVal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    if (TryParseAsHexLiteral(valueToken, out int outputHex))
                    {
                        result = outputHex.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out int outputOctal))
                    {
                        result = outputOctal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }
                }
                else if (conformTypeName.Equals(Tokens.LongLong))
                {
                    if (LetCoercer.TryCoerce(valueToken, out long newVal))
                    {
                        result = newVal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    if (TryParseAsHexLiteral(valueToken, out long outputHex))
                    {
                        result = outputHex.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out long outputOctal))
                    {
                        result = outputOctal.ToString(CultureInfo.InvariantCulture);
                        return true;
                    }
                }

                if (isArabicNumber)
                {
                    //If we get here the type cannot conform because it is an overflow
                    result = checkValue.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
            }

            else if (conformTypeName.Equals(Tokens.Double) || conformTypeName.Equals(Tokens.Single)
                || conformTypeName.Equals(Tokens.Currency))
            {
                //If a coercion fails, it may be an overflow condition for the conformTypeName
                var isRational = LetCoercer.TryCoerce(valueToken, out double checkValue);

                var derivedTypeName = DeriveTypeName(valueToken, out bool usedTypehint);
                if (derivedTypeName.Equals(Tokens.Date))
                {
                    if (LetCoercer.TryCoerce(valueToken, out ComparableDateValue dvComparable))
                    {
                        result = Convert.ToDouble(dvComparable.AsDecimal).ToString(CultureInfo.InvariantCulture);
                        return true;
                    }
                }
                else if(conformTypeName.Equals(Tokens.Single) && LetCoercer.TryCoerce(valueToken, out float floatVal))
                {
                    result = floatVal.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
                else if (conformTypeName.Equals(Tokens.Double) && LetCoercer.TryCoerce(valueToken, out double newVal))
                {
                    result = newVal.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
                else if (conformTypeName.Equals(Tokens.Currency) && LetCoercer.TryCoerce(valueToken, out decimal decimalVal))
                {
                    result = decimalVal.ToString(CultureInfo.InvariantCulture);
                    return true;
                }

                if (isRational)
                {
                    //If we get here the type cannot conform because it is an overflow
                    result = checkValue.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
            }
            else if (conformTypeName.Equals(Tokens.Boolean))
            {
                if (LetCoercer.TryCoerce(valueToken, out bool newVal))
                {
                    result = newVal.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
            }
            else if (conformTypeName.Equals(Tokens.String))
            {
                result = valueToken;
                return IsStringConstant(valueToken);
            }
            else if (conformTypeName.Equals(Tokens.Date))
            {
                var derivedTypeName = DeriveTypeName(valueToken, out bool _);
                if (!derivedTypeName.Equals(Tokens.Date))
                {
                    var sourceType = derivedTypeName.Equals(string.Empty) ? Tokens.String : derivedTypeName;
                    var (Success, CoercedText) = LetCoerce((sourceType, valueToken), Tokens.Date);
                    result = CoercedText;
                    return Success;
                }
                else if (LetCoercer.TryCoerce(valueToken, out ComparableDateValue dvComparable))
                {
                    valueToken = dvComparable.AsDate.ToString(CultureInfo.InvariantCulture);
                    result = dvComparable.AsDateLiteral();
                    return true;
                }
            }
            return false;
        }

        private static Dictionary<string, Dictionary<string, Func<string, (bool, string)>>> LetCoercions = new Dictionary<string, Dictionary<string, Func<string, (bool, string)>>>()
        {
            [Tokens.String] = new Dictionary<string, Func<string, (bool, string)>>
            {
                [Tokens.String] = (a) => { return (true, a); },
                [Tokens.Date] = StringToDate,
            },
        };

        private static (bool, string) StringToDate(string sourceText)
        {
            if (LetCoercer.TryCoerce(sourceText, out ComparableDateValue dvComparable))
            {
                return (true, dvComparable.AsDateLiteral());
            }
            return (false, sourceText);
        }


        private static (bool Success, string CoercedText) LetCoerce((string Type, string Text) source, string destinationType)
        {
            var returnValue = LetCoercer.TryCoerceToken(source, destinationType, out string coercedValue);
            return (returnValue, coercedValue);
        }

        private static bool TryParseAsDateLiteral(string valueString, out ComparableDateValue value)
        {
            value = default;
            if (IsDateLiteral(valueString))
            {
                return LetCoercer.TryCoerce(valueString, out value);
            }
            return false;
        }

        private static string[] HexPrefixes = new string[] { "&h", "&H" };

        private static bool TryParseAsHexLiteral(string valueString, out short value)
        {
            (bool success, short result) = TryParseAsHexLiteral(valueString, HexPrefixes, Convert.ToInt16);
            value = result;
            return success;
        }

        private static bool TryParseAsHexLiteral(string valueString, out int value)
        {
            (bool success, int result) = TryParseAsHexLiteral(valueString, HexPrefixes, Convert.ToInt32);
            value = result;
            return success;
        }

        private static bool TryParseAsHexLiteral(string valueString, out long value)
        {
            (bool success, long result) = TryParseAsHexLiteral(valueString, HexPrefixes, Convert.ToInt64);
            value = result;
            return success;
        }

        private static (bool, T) TryParseAsHexLiteral<T>(string valueString, string[] prefixs, Func<string, int, T> conversion)
        {
            T value = default;

            if (!prefixs.Any(pf => valueString.StartsWith(pf)))
            {
                return (false, value);
            }

            var hexString = valueString.Substring(2).ToUpperInvariant();
            try
            {
                value = conversion(hexString, 16);
                return (true, value);
            }
            catch (OverflowException)
            {
                return (false, value);
            }
        }

        private static string[] OctalPrefixes = new string[] { "&o", "&O" };

        private static bool TryParseAsOctalLiteral(string valueString, out short value)
        {
            (bool success, short result) = TryParseAsOctalLiteral(valueString, OctalPrefixes, Convert.ToInt16);
            value = result;
            return success;
        }

        private static bool TryParseAsOctalLiteral(string valueString, out int value)
        {
            (bool success, int result) = TryParseAsOctalLiteral(valueString, OctalPrefixes, Convert.ToInt32);
            value = result;
            return success;
        }

        private static bool TryParseAsOctalLiteral(string valueString, out long value)
        {
            (bool success, long result) = TryParseAsOctalLiteral(valueString, OctalPrefixes, Convert.ToInt64);
            value = result;
            return success;
        }

        private static (bool, T) TryParseAsOctalLiteral<T>(string valueString, string[] prefixs, Func<string, int, T> conversion)
        {
            T value = default;

            if (!prefixs.Any(pf => valueString.StartsWith(pf)))
            {
                return (false, value);
            }

            var octalString = valueString.Substring(2);
            try
            {
                value = conversion(octalString, 8);
                return (true, value);
            }
            catch (OverflowException)
            {
                return (false, value);
            }
        }

        private static bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");

        private static bool IsDateLiteral(string input) => input.StartsWith("#") && input.EndsWith("#");
    }
}
