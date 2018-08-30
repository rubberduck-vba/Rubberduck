using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public struct VBACurrency
    {
        public static decimal MinValue = -922337203685477.5808M;
        public static decimal MaxValue = 922337203685477.5807M;
        public static decimal Parse(string valueText)
        {
            var checkValue = Math.Round(decimal.Parse(valueText), 4, MidpointRounding.ToEven);
            return MinValue < checkValue && MaxValue > checkValue ? checkValue 
                : throw new OverflowException();
        }

        public static bool TryParse(string valueText, out decimal value)
        {
            value = default;
            try
            {
                value = Parse(valueText);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }
        }
    }

    public struct LetCoercer
    {
        //Content: Dictionary<sourceTypeName,Dictionary<LetDestinationTypeName,CoercionFunc>
        private static Dictionary<string, Dictionary<string, Func<string, string>>> _coercions;

        public static bool TryCoerceToken((string Type, string Text) source, string destinationType, out string resultToken)
        {
            resultToken = string.Empty;
            try
            {
                resultToken = CoerceToken(source, destinationType);
                return true;
            }
            catch(ArgumentNullException)
            {
                return false;
            }
            catch (OverflowException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (KeyNotFoundException knf)
            {
#if DEBUG
                throw knf;
#else
                return false;
#endif
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static string CoerceToken((string Type, string Text) source, string destinationType)
        {
            InitializeCoercions();

            if (!_coercions.ContainsKey(source.Type))
            {
                throw new KeyNotFoundException($"Let Coercion source type: {source.Type} not supported");
            }
            if (!_coercions[source.Type].ContainsKey(destinationType))
            {
                throw new KeyNotFoundException($"Let Coercion source=>destination pair: {source.Type}=>{destinationType} not supported");
            }

            return _coercions[source.Type][destinationType](source.Text);
        }

        public static bool TryCoerce(string valueText, out byte value)
            => TryCoerce(valueText, Tokens.Byte, out value, byte.Parse);

        public static bool TryCoerce(string valueText, out Int16 value)
            => TryCoerce(valueText, Tokens.Integer, out value, Int16.Parse);

        public static bool TryCoerce(string valueText, out Int32 value)
            => TryCoerce(valueText, Tokens.Long, out value, Int32.Parse);

        public static bool TryCoerce(string valueText, out long value)
            => TryCoerce(valueText, Tokens.LongLong, out value, long.Parse);

        public static bool TryCoerce(string valueText, out double value)
            => TryCoerce(valueText, Tokens.Double, out value, double.Parse);

        public static bool TryCoerce(string valueText, out float value)
            => TryCoerce(valueText, Tokens.Single, out value, float.Parse);

        public static bool TryCoerce(string valueText, out decimal value)
            => TryCoerce(valueText, Tokens.Currency, out value, VBACurrency.Parse);

        public static bool TryCoerce(string valueText, out bool value)
            => TryCoerce(valueText, Tokens.Boolean, out value, bool.Parse);

        public static bool TryCoerce(string valueText, out string value)
        {
            value = valueText;
            return true;
        }

        public static bool TryCoerce(string valueText, out ComparableDateValue value)
            => TryCoerce(valueText, Tokens.Date, out value, ComparableDateValue.Parse);

        public static bool ExceedsValueTypeRange(string valueType, string token)
        {
            try
            {
                CoerceToken((Tokens.String, token), valueType);
            }
            catch (OverflowException)
            {
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (ArgumentNullException)
            {
                return false;
            }
            return false;
        }

        private static bool TryCoerce<T>(string valueText, string typeName, out T value, Func<string, T> parser)
        {
            value = default;
            if (TryCoerceToken((Tokens.String, valueText), typeName, out string valueToken))
            {
                value = parser(valueToken);
                return true;
            }
            return false;
        }

        private static void InitializeCoercions()
        {
            if (_coercions != null) { return; }

            _coercions = new Dictionary<string, Dictionary<string, Func<string, string>>>
            {
                [Tokens.String] = new Dictionary<string, Func<string,string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { return byte.Parse(StringToByte(a)).ToString(); },
                    [Tokens.Integer] = (a) => { return short.Parse(StringToIntegral(a)).ToString(); },
                    [Tokens.Long] = (a) => { return Int32.Parse(StringToIntegral(a)).ToString(); },
                    [Tokens.LongLong] = (a) => { return Int64.Parse(StringToIntegral(a)).ToString(); },
                    [Tokens.Double] = (a) => { return double.Parse(StringToRational(a)).ToString(); },
                    [Tokens.Single] = (a) => { return float.Parse(StringToRational(a)).ToString(); },
                    [Tokens.Currency] = (a) => { return VBACurrency.Parse(StringToRational(a)).ToString(); },
                    [Tokens.Boolean] = (a) => { return bool.Parse(StringToBoolean(a)) ? Tokens.True : Tokens.False; },
                    [Tokens.Date] = (a) => { return StringToDate(a); }
                },

                [Tokens.Byte] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { return a; },
                    [Tokens.Integer] = (a) => { return a; },
                    [Tokens.Long] = (a) => { return a; },
                    [Tokens.LongLong] = (a) => { return a; },
                    [Tokens.Double] = (a) => { return a; },
                    [Tokens.Single] = (a) => { return a; },
                    [Tokens.Currency] = (a) => { return a; },
                    [Tokens.Boolean] = (a) => { return NumericToBoolean(a); },
                    [Tokens.Date] = (a) => { return NumericToDate(a); },
                },

                [Tokens.Integer] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { return byte.Parse(a).ToString(); },
                    [Tokens.Integer] = (a) => { return a; },
                    [Tokens.Long] = (a) => { return a; },
                    [Tokens.LongLong] = (a) => { return a; },
                    [Tokens.Double] = (a) => { return a; },
                    [Tokens.Single] = (a) => { return a; },
                    [Tokens.Currency] = (a) => { return a; },
                    [Tokens.Boolean] = (a) => { return NumericToBoolean(a); },
                    [Tokens.Date] = (a) => { return NumericToDate(a); },
                },

                [Tokens.Long] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { return byte.Parse(a).ToString(); },
                    [Tokens.Integer] = (a) => { return Int16.Parse(a).ToString(); },
                    [Tokens.Long] = (a) => { return a; },
                    [Tokens.LongLong] = (a) => { return a; },
                    [Tokens.Double] = (a) => { return a; },
                    [Tokens.Single] = (a) => { return a; },
                    [Tokens.Currency] = (a) => { return a; },
                    [Tokens.Boolean] = (a) => { return NumericToBoolean(a); },
                    [Tokens.Date] = (a) => { return NumericToDate(a); },
                },

                [Tokens.Double] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { a = BankersRound(a); return byte.Parse(a).ToString(); },
                    [Tokens.Integer] = (a) => { a = BankersRound(a); return Int16.Parse(a).ToString(); },
                    [Tokens.Long] = (a) => { a = BankersRound(a); return Int32.Parse(a).ToString(); },
                    [Tokens.LongLong] = (a) => { a = BankersRound(a); return long.Parse(a).ToString(); },
                    [Tokens.Double] = (a) => { return a; },
                    [Tokens.Single] = (a) => { return float.Parse(a).ToString(); },
                    [Tokens.Currency] = (a) => { return decimal.Parse(a).ToString(); },
                    [Tokens.Boolean] = (a) => { return NumericToBoolean(a); },
                    [Tokens.Date] = (a) => { return NumericToDate(a); },
                },

                [Tokens.Single] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { a = BankersRound(a); return byte.Parse(a).ToString(); },
                    [Tokens.Integer] = (a) => { a = BankersRound(a); return Int16.Parse(a).ToString(); },
                    [Tokens.Long] = (a) => { a = BankersRound(a); return Int32.Parse(a).ToString(); },
                    [Tokens.LongLong] = (a) => { a = BankersRound(a); return long.Parse(a).ToString(); },
                    [Tokens.Double] = (a) => { return a; },
                    [Tokens.Single] = (a) => { return a; },
                    [Tokens.Currency] = (a) => { return decimal.Parse(a).ToString(); },
                    [Tokens.Boolean] = (a) => { return NumericToBoolean(a); },
                    [Tokens.Date] = (a) => { return NumericToDate(a); },
                },

                [Tokens.Currency] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return a; },
                    [Tokens.Byte] = (a) => { a = BankersRound(a); return byte.Parse(a).ToString(); },
                    [Tokens.Integer] = (a) => { a = BankersRound(a); return Int16.Parse(a).ToString(); },
                    [Tokens.Long] = (a) => { a = BankersRound(a); return Int32.Parse(a).ToString(); },
                    [Tokens.LongLong] = (a) => { a = BankersRound(a); return long.Parse(a).ToString(); },
                    [Tokens.Double] = (a) => { return a; },
                    [Tokens.Single] = (a) => { return float.Parse(a).ToString(); },
                    [Tokens.Currency] = (a) => { return a; },
                    [Tokens.Boolean] = (a) => { return NumericToBoolean(a); },
                    [Tokens.Date] = (a) => { return NumericToDate(a); },
                },

                [Tokens.Boolean] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return BooleanToString(a); },
                    [Tokens.Byte] = (a) => { var val = BooleanAsLong(a); return val != 0 ? byte.MaxValue.ToString() : byte.MinValue.ToString(); },
                    [Tokens.Integer] = (a) => { return BooleanAsLong(a).ToString(); },
                    [Tokens.Long] = (a) => { return BooleanAsLong(a).ToString(); },
                    [Tokens.LongLong] = (a) => { return BooleanAsLong(a).ToString(); },
                    [Tokens.Double] = (a) => { return BooleanAsLong(a).ToString(); },
                    [Tokens.Single] = (a) => { return BooleanAsLong(a).ToString(); },
                    [Tokens.Currency] = (a) => { return BooleanAsLong(a).ToString(); },
                    [Tokens.Boolean] = (a) => { return a; },
                    [Tokens.Date] = (a) => { var val = BooleanAsLong(a); return NumericToDate(val.ToString()); },
                },

                [Tokens.Date] = new Dictionary<string, Func<string, string>>
                {
                    [Tokens.String] = (a) => { return RemoveStartAndEnd(a, "#"); },
                    [Tokens.Byte] = (a) => { var result = ComparableDateValue.Parse(a); return Convert.ToByte(result.AsDecimal).ToString(); },
                    [Tokens.Integer] = (a) => { var result = ComparableDateValue.Parse(a); return Convert.ToInt16(result.AsDecimal).ToString(); },
                    [Tokens.Long] = (a) => { var result = ComparableDateValue.Parse(a); return Convert.ToInt32(result.AsDecimal).ToString(); },
                    [Tokens.LongLong] = (a) => { var result = ComparableDateValue.Parse(a); return Convert.ToInt64(result.AsDecimal).ToString(); },
                    [Tokens.Double] = (a) => { var result = ComparableDateValue.Parse(a); return Convert.ToDouble(result.AsDecimal).ToString(); },
                    [Tokens.Single] = (a) => { var result = ComparableDateValue.Parse(a); return float.Parse(result.AsDecimal.ToString()).ToString(); },
                    [Tokens.Currency] = (a) => { var result = ComparableDateValue.Parse(a); return result.AsDecimal.ToString(); },
                    [Tokens.Boolean] = (a) => { var result = ComparableDateValue.Parse(a); return result.AsDecimal != 0 ? Tokens.True : Tokens.False; },
                    [Tokens.Date] = (a) => { return a; },
                },
            };
        }

        private static string StringToDate(string sourceText)
        {
            int? intValue = BooleanTokenToInt(sourceText);
            var parseValue = intValue.HasValue ? intValue.ToString() : sourceText;

            if (double.TryParse(parseValue, out double doubleValue))
            {
                return NumericToDate(parseValue);
            }
            if (ComparableDateValue.TryParse(AnnotateAsDateLiteral(parseValue), out ComparableDateValue dvComparable))
            {
                return dvComparable.AsDateLiteral();
            }
            throw new FormatException();
        }

        private static string StringToByte(string sourceText)
        {
            int? intValue = BooleanTokenToInt(sourceText);
            if (intValue.HasValue)
            {
                return intValue == 0 ? byte.MinValue.ToString() : byte.MaxValue.ToString();
            }
            return StringToIntegral(sourceText);
        }

        private static string StringToIntegral(string sourceText)
        {
            return BankersRound(StringToRational(sourceText));
        }

        private static string StringToRational(string sourceText)
        {
            sourceText = RemoveDoubleQuotes(sourceText);

            int? intValue = BooleanTokenToInt(sourceText);
            var parseValue = intValue.HasValue ? intValue.ToString() : sourceText;

            return decimal.TryParse(parseValue, out decimal decValue) ? decValue.ToString()
                : double.TryParse(parseValue, out double value) ? value.ToString() : sourceText;
        }

        private static string StringToBoolean(string sourceText)
        {
            sourceText = RemoveDoubleQuotes(sourceText);
            if (sourceText.Equals(Tokens.True, StringComparison.OrdinalIgnoreCase))
            {
                return Tokens.True;
            }
            if (sourceText.Equals(Tokens.False, StringComparison.OrdinalIgnoreCase))
            {
                return Tokens.False;
            }
            if (sourceText.Equals($"#{Tokens.True}#", StringComparison.Ordinal))
            {
                return Tokens.True;
            }
            if (sourceText.Equals($"#{Tokens.False}#", StringComparison.Ordinal))
            {
                return Tokens.False;
            }
            if (double.TryParse(sourceText, out double asDouble))
            {
                return asDouble != 0 ? Tokens.True : Tokens.False;
            }
            return string.Empty;
        }

        private static string NumericToDate(string source)
        {
            if (double.TryParse(source, out double dateAsDouble))
            {
                var dv = new DateValue(DateTime.FromOADate(dateAsDouble));
                var dateValue = new ComparableDateValue(dv);
                return dateValue.AsDateLiteral();
            }
            return string.Empty;
        }

        private static int? BooleanTokenToInt(string sourceText)
        {
            if (sourceText.Equals(Tokens.True, StringComparison.OrdinalIgnoreCase))
            {
                return -1;
            }
            if (sourceText.Equals(Tokens.False, StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }
            if (sourceText.Equals($"#{Tokens.True}#", StringComparison.Ordinal))
            {
                return -1;
            }
            if (sourceText.Equals($"#{Tokens.False}#", StringComparison.Ordinal))
            {
                return 0;
            }
            return null;
        }

        private static string RemoveDoubleQuotes(string source)
        {
            return RemoveStartAndEnd(source, "\"");
        }

        private static string RemoveStartAndEnd(string source, string character)
        {
            string result = source;
            if (result.StartsWith(character))
            {
                result = result.Remove(0, 1);
            }
            if (result.EndsWith(character))
            {
                result = result.Remove(result.Length - 1);
            }
            return result;
        }

        private static string BooleanToString(string source)
        {
            if (source.Equals(Tokens.True) || source.Equals(Tokens.False))
            {
                return source;
            }

            return double.Parse(source) != 0 ? Tokens.True : Tokens.False;
        }

        private static string BankersRound(string source)
             => Math.Round(double.Parse(source), MidpointRounding.ToEven).ToString();

        private static string NumericToBoolean(string source)
             => double.Parse(source) != 0 ? Tokens.True : Tokens.False;

        private static long BooleanAsLong(string source)
            => source.Equals(Tokens.True) ? -1 : 0;

        private static string AnnotateAsDateLiteral(string input)
        {
            var result = input;
            if (!input.StartsWith("#"))
            {
                result = $"#{result}";
            }
            if (!input.EndsWith("#"))
            {
                result = $"{result}#";
            }
            return result;
        }
    }
}
