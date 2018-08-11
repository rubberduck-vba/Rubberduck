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
            result = valueToken;
            if (valueToken is null || conformTypeName is null)
            {
                return false;
            }

            if (valueToken.Equals(double.NaN.ToString(CultureInfo.InvariantCulture)) &&
                !conformTypeName.Equals(Tokens.String))
            {
                result = string.Empty;
                return false;
            }

            else if (conformTypeName.Equals(Tokens.LongLong) || conformTypeName.Equals(Tokens.Long) ||
                conformTypeName.Equals(Tokens.Integer) || conformTypeName.Equals(Tokens.Byte))
            {
                if (TokenParser.TryParse(valueToken, out long newVal))
                {
                    result = newVal.ToString();
                    return true;
                }

                if (conformTypeName.Equals(Tokens.Integer))
                {
                    if (TryParseAsHexLiteral(valueToken, out short outputHex))
                    {
                        result = outputHex.ToString();
                        return true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out short outputOctal))
                    {
                        result = outputOctal.ToString();
                        return true;
                    }
                }

                if (conformTypeName.Equals(Tokens.Long))
                {
                    if (TryParseAsHexLiteral(valueToken, out int outputHex))
                    {
                        result = outputHex.ToString();
                        return true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out int outputOctal))
                    {
                        result = outputOctal.ToString();
                        return true;
                    }
                }

                if (conformTypeName.Equals(Tokens.LongLong))
                {
                    if (TryParseAsHexLiteral(valueToken, out long outputHex))
                    {
                        result = outputHex.ToString();
                        return true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out long outputOctal))
                    {
                        result = outputOctal.ToString();
                        return true;
                    }
                }
            }
            else if (conformTypeName.Equals(Tokens.Double) || conformTypeName.Equals(Tokens.Single))
            {
                var derivedTypeName = DeriveTypeName(valueToken, out bool usedTypehint);
                if (derivedTypeName.Equals(Tokens.Date))
                {
                    if (TokenParser.TryParse(valueToken, out ComparableDateValue dvComparable))
                    {
                        result = Convert.ToDouble(dvComparable.AsDecimal).ToString(CultureInfo.InvariantCulture);
                        return true;
                    }
                }
                else if (TokenParser.TryParse(valueToken, out double newVal))
                {
                    result = newVal.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
            }
            else if (conformTypeName.Equals(Tokens.Boolean))
            {
                if (TokenParser.TryParse(valueToken, out bool newVal))
                {
                    result = newVal.ToString();
                    return true;
                }
            }
            else if (conformTypeName.Equals(Tokens.String))
            {
                return IsStringConstant(valueToken);
            }
            else if (conformTypeName.Equals(Tokens.Currency))
            {
                if (TokenParser.TryParse(valueToken, out decimal newVal))
                {
                    var currencyValue = Math.Round(newVal, 4, MidpointRounding.ToEven);
                    result = currencyValue.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
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
                else if (TokenParser.TryParse(valueToken, out ComparableDateValue dvComparable))
                {
                    valueToken = dvComparable.AsDate.ToString(CultureInfo.InvariantCulture);
                    result = ComparableDateValue.AsDateLiteral(valueToken);
                    return true;
                }
            }
            return false;
        }

        private static string DeriveTypeName(string inputString, out bool derivedFromTypeHint)
        {
            if (TryDeriveTypeName(inputString, out (string TypeName, string Value) result, out derivedFromTypeHint))
            {
                return result.TypeName;
            }
            return string.Empty;
        }

        private static (bool, string) LetCoerce((string Type, string Text) source, string destinationType)
        {
            var returnValue = LetCoercer.TryLetCoerce(source, destinationType, out string coercedValue);
            return (returnValue, coercedValue);
        }

        private static bool TryParseAsDateLiteral(string valueString, out ComparableDateValue value)
        {
            return TokenParser.TryParse(valueString, out value);
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
    }
}
