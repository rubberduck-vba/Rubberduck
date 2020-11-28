using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Refactoring.ParseTreeValue
{
    public struct TypeTokenPair
    {
        public string ValueType { get; }
        public string Token { get; }

        public TypeTokenPair(string valueType, string token)
        {
            ValueType = valueType;
            Token = token;
        }

        public bool HasValue => Token != null;
        public bool HasType => ValueType != null;

        public static bool TryParse(string inputString, out TypeTokenPair result)
        {
            result = new TypeTokenPair(null, inputString);
            if (inputString is null || inputString.Length == 0)
            {
                return false;
            }

            if (IsDateLiteral(inputString))
            {
                result = ConformToType(Tokens.Date, inputString, false);
                if (result.HasValue)
                {
                    return true;
                }
            }

            if (inputString.StartsWith("\"") && inputString.EndsWith("\""))
            {
                result = ConformToType(Tokens.String, inputString, false);
                return true;
            }

            if (inputString.Equals(Tokens.True) || inputString.Equals(Tokens.False))
            {
                result = ConformToType(Tokens.Boolean, inputString, false);
                return true;
            }

            if (inputString.Contains(".") || inputString.Count(ch => ch.Equals('E')) == 1)
            {
                result = ConformToType(Tokens.Double, inputString, false);
                if (result.HasValue)
                {
                    return true;
                }
            }

            result = ConformToType(Tokens.Integer, inputString, false);
            if (result.HasValue)
            {
                return true;
            }

            result = ConformToType(Tokens.Long, inputString, false);
            if (result.HasValue)
            {
                return true;
            }

            result = ConformToType(Tokens.LongLong, inputString, false);
            if (result.HasValue)
            {
                result = new TypeTokenPair(Tokens.Double, result.Token);
                return true;
            }

            return false;
        }

        public static TypeTokenPair ConformToType(string goalValueType, string valueToken, bool allowOverflow = true)
        {
            if (valueToken.Equals(double.NaN.ToString(CultureInfo.InvariantCulture)) &&
                !goalValueType.Equals(Tokens.String))
            {
                return new TypeTokenPair(goalValueType, null);
            }

            if (conformToken.ContainsKey(goalValueType))
            {
                return conformToken[goalValueType](valueToken, allowOverflow);
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static Dictionary<string, Func<string, bool, TypeTokenPair>> conformToken = new Dictionary<string, Func<string, bool, TypeTokenPair>>()
        {
            [Tokens.Boolean] = ConformTokenToBoolean,
            [Tokens.Byte] = ConformTokenToByte,
            [Tokens.Integer] = ConformTokenToInteger,
            [Tokens.Long] = ConformTokenToLong,
            [Tokens.LongLong] = ConformTokenToLongLong,
            [Tokens.Single] = ConformTokenToSingle,
            [Tokens.Double] = ConformTokenToDouble,
            [Tokens.Currency] = ConformTokenToCurrency,
            [Tokens.String] = ConformTokenToString,
            [Tokens.Date] = ConformTokenToDate,
        };


        private static TypeTokenPair ConformTokenToInteger(string valueToken, bool allowOverflow = true)
        {
            var goalValueType = Tokens.Integer;
            if (LetCoerce.TryCoerce(valueToken, out short newVal)
                || TryParseAsHexLiteral(valueToken, out newVal)
                || TryParseAsOctalLiteral(valueToken, out newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            if (allowOverflow && LetCoerce.TryCoerce(valueToken, out long overflowValue))
            {
                return new TypeTokenPair(goalValueType, overflowValue.ToString(CultureInfo.InvariantCulture));
            }

            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToLong(string valueToken, bool allowOverflow = true)
        {
            var goalValueType = Tokens.Long;
            if (LetCoerce.TryCoerce(valueToken, out int newVal)
                || TryParseAsHexLiteral(valueToken, out newVal)
                || TryParseAsOctalLiteral(valueToken, out newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            if (allowOverflow && LetCoerce.TryCoerce(valueToken, out long overflowValue))
            {
                return new TypeTokenPair(goalValueType, overflowValue.ToString(CultureInfo.InvariantCulture));
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToLongLong(string valueToken, bool allowOverflow = false)
        {
            var goalValueType = Tokens.LongLong;
            if (LetCoerce.TryCoerce(valueToken, out long newVal)
                || TryParseAsHexLiteral(valueToken, out newVal)
                || TryParseAsOctalLiteral(valueToken, out newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToByte(string valueToken, bool allowOverflow = true)
        {
            var goalValueType = Tokens.Byte;
            if (LetCoerce.TryCoerce(valueToken, out byte newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            if (allowOverflow && LetCoerce.TryCoerce(valueToken, out long overflowValue))
            {
                return new TypeTokenPair(goalValueType, overflowValue.ToString(CultureInfo.InvariantCulture));
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToBoolean(string valueToken, bool allowOverflow = false)
        {
            var goalValueType = Tokens.Boolean;
            if (LetCoerce.TryCoerce(valueToken, out bool newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToSingle(string valueToken, bool allowOverflow = true)
        {
            var goalValueType = Tokens.Single;
            if (LetCoerce.TryCoerce(valueToken, out float newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            if (allowOverflow && LetCoerce.TryCoerce(valueToken, out double overflowValue))
            {
                return new TypeTokenPair(goalValueType, overflowValue.ToString(CultureInfo.InvariantCulture));
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToDouble(string valueToken, bool allowOverflow = false)
        {
            var goalValueType = Tokens.Double;
            if (LetCoerce.TryCoerce(valueToken, out double newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToCurrency(string valueToken, bool allowOverflow = true)
        {
            var goalValueType = Tokens.Currency;
            if (LetCoerce.TryCoerce(valueToken, out decimal newVal))
            {
                return new TypeTokenPair(goalValueType, newVal.ToString(CultureInfo.InvariantCulture));
            }

            if (allowOverflow && LetCoerce.TryCoerce(valueToken, out double overflowValue))
            {
                return new TypeTokenPair(goalValueType, overflowValue.ToString(CultureInfo.InvariantCulture));
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToDate(string valueToken, bool allowOverflow = false)
        {
            var goalValueType = Tokens.Date;
            if (LetCoerce.TryCoerce(valueToken, out ComparableDateValue dvComparable))
            {
                return new TypeTokenPair(goalValueType, dvComparable.AsDateLiteral());
            }
            return new TypeTokenPair(goalValueType, null);
        }

        private static TypeTokenPair ConformTokenToString(string valueToken, bool allowOverflow = false)
            => new TypeTokenPair(Tokens.String, valueToken);

        private static bool TryParseAsDateLiteral(string valueString, out ComparableDateValue value)
        {
            value = default;
            if (IsDateLiteral(valueString))
            {
                return LetCoerce.TryCoerce(valueString, out value);
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

        private static bool IsDateLiteral(string input) => input.StartsWith("#") && input.EndsWith("#");
    }
}
