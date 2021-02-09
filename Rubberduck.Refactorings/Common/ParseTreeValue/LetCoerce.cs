using System;
using System.Collections.Generic;
using System.Globalization;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;

namespace Rubberduck.Refactoring.ParseTreeValue
{
    public struct VBACurrency
    {
        public static decimal MinValue = -922337203685477.5808M;
        public static decimal MaxValue = 922337203685477.5807M;
        public static decimal Parse(string valueText)
        {
            var checkValue = Math.Round(decimal.Parse(valueText, NumberStyles.Float, CultureInfo.InvariantCulture), 4, MidpointRounding.ToEven);
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

    public struct LetCoerce
    {
        //Content: Dictionary<sourceTypeName,Dictionary<LetDestinationTypeName,CoercionFunc>
        private static Dictionary<string, Dictionary<string, Func<string, string>>> _coercions = new Dictionary<string, Dictionary<string, Func<string, string>>>
        {
            [Tokens.String] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => byte.Parse(StringToByte(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => short.Parse(StringToIntegral(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => int.Parse(StringToIntegral(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.LongLong] = a => long.Parse(StringToIntegral(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Double] = a => double.Parse(StringToRational(a), NumberStyles.Float, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Single] = a => float.Parse(StringToRational(a), NumberStyles.Float, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Currency] = a => VBACurrency.Parse(StringToRational(a)).ToString(CultureInfo.InvariantCulture),
                [Tokens.Boolean] = a => bool.Parse(StringToBoolean(a)) ? Tokens.True : Tokens.False,
                [Tokens.Date] = StringToDate
            },

            [Tokens.Byte] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => a,
                [Tokens.Integer] = a => a,
                [Tokens.Long] = a => a,
                [Tokens.LongLong] = a => a,
                [Tokens.Double] = a => a,
                [Tokens.Single] = a => a,
                [Tokens.Currency] = a => a,
                [Tokens.Boolean] = NumericToBoolean,
                [Tokens.Date] = NumericToDate,
            },

            [Tokens.Integer] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => byte.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => a,
                [Tokens.Long] = a => a,
                [Tokens.LongLong] = a => a,
                [Tokens.Double] = a => a,
                [Tokens.Single] = a => a,
                [Tokens.Currency] = a => a,
                [Tokens.Boolean] = NumericToBoolean,
                [Tokens.Date] = NumericToDate,
            },

            [Tokens.Long] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => byte.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => short.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => a,
                [Tokens.LongLong] = a => a,
                [Tokens.Double] = a => a,
                [Tokens.Single] = a => a,
                [Tokens.Currency] = a => a,
                [Tokens.Boolean] = NumericToBoolean,
                [Tokens.Date] = NumericToDate,
            },

            [Tokens.Double] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => byte.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => short.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => int.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.LongLong] = a => long.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Double] = a => a,
                [Tokens.Single] = a => float.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Currency] = a => decimal.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Boolean] = NumericToBoolean,
                [Tokens.Date] = NumericToDate,
            },

            [Tokens.Single] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => byte.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => short.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => int.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.LongLong] = a => long.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Double] = a => a,
                [Tokens.Single] = a => a,
                [Tokens.Currency] = a => decimal.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Boolean] = NumericToBoolean,
                [Tokens.Date] = NumericToDate,
            },

            [Tokens.Currency] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => a,
                [Tokens.Byte] = a => byte.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => short.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => int.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.LongLong] = a => long.Parse(BankersRound(a), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Double] = a => a,
                [Tokens.Single] = a => float.Parse(a, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Currency] = a => a,
                [Tokens.Boolean] = NumericToBoolean,
                [Tokens.Date] = NumericToDate,
            },

            [Tokens.Boolean] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = BooleanToString,
                [Tokens.Byte] = a => BooleanAsLong(a) != 0 ? byte.MaxValue.ToString(CultureInfo.InvariantCulture) : byte.MinValue.ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => BooleanAsLong(a).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => BooleanAsLong(a).ToString(CultureInfo.InvariantCulture),
                [Tokens.LongLong] = a => BooleanAsLong(a).ToString(CultureInfo.InvariantCulture),
                [Tokens.Double] = a => BooleanAsLong(a).ToString(CultureInfo.InvariantCulture),
                [Tokens.Single] = a => BooleanAsLong(a).ToString(CultureInfo.InvariantCulture),
                [Tokens.Currency] = a => BooleanAsLong(a).ToString(CultureInfo.InvariantCulture),
                [Tokens.Boolean] = a => a,
                [Tokens.Date] = a => NumericToDate(BooleanAsLong(a).ToString(CultureInfo.InvariantCulture)),
            },

            [Tokens.Date] = new Dictionary<string, Func<string, string>>
            {
                [Tokens.String] = a => RemoveStartAndEnd(a, "#"),
                [Tokens.Byte] = a => Convert.ToByte(ComparableDateValue.Parse(a).AsDecimal, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Integer] = a => Convert.ToInt16(ComparableDateValue.Parse(a).AsDecimal, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Long] = a => Convert.ToInt32(ComparableDateValue.Parse(a).AsDecimal, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.LongLong] = a => Convert.ToInt64(ComparableDateValue.Parse(a).AsDecimal, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Double] = a => Convert.ToDouble(ComparableDateValue.Parse(a).AsDecimal, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Single] = a => float.Parse(ComparableDateValue.Parse(a).AsDecimal.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture),
                [Tokens.Currency] = a => ComparableDateValue.Parse(a).AsDecimal.ToString(CultureInfo.InvariantCulture),
                [Tokens.Boolean] = a => ComparableDateValue.Parse(a).AsDecimal != 0 ? Tokens.True : Tokens.False,
                [Tokens.Date] = a => a,
            },
        };


        public static bool TryCoerceToken((string Type, string Text) source, string destinationType, out string resultToken)
        {
            resultToken = string.Empty;
            try
            {
                resultToken = Coerce(source, destinationType);
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

        public static string Coerce((string Type, string Text) source, string destinationType)
        {
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
            => TryCoerce(valueText, Tokens.Byte, out value, s => byte.Parse(s, CultureInfo.InvariantCulture));

        public static bool TryCoerce(string valueText, out short value)
            => TryCoerce(valueText, Tokens.Integer, out value, s => short.Parse(s, CultureInfo.InvariantCulture));

        public static bool TryCoerce(string valueText, out int value)
            => TryCoerce(valueText, Tokens.Long, out value, s => int.Parse(s, CultureInfo.InvariantCulture));

        public static bool TryCoerce(string valueText, out long value)
            => TryCoerce(valueText, Tokens.LongLong, out value, s => long.Parse(s, CultureInfo.InvariantCulture));

        public static bool TryCoerce(string valueText, out double value)
            => TryCoerce(valueText, Tokens.Double, out value, s => double.Parse(s, CultureInfo.InvariantCulture));

        public static bool TryCoerce(string valueText, out float value)
            => TryCoerce(valueText, Tokens.Single, out value, s => float.Parse(s, CultureInfo.InvariantCulture));

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
                Coerce((Tokens.String, token), valueType);
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
            if (TryCoerceToken((Tokens.String, valueText), typeName, out var valueToken))
            {
                value = parser(valueToken);
                return true;
            }
            return false;
        }

        private static string StringToDate(string sourceText)
        {
            var intValue = BooleanTokenToInt(sourceText);
            var parseValue = intValue.HasValue ? intValue.ToString() : sourceText;

            if (double.TryParse(parseValue, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                return NumericToDate(parseValue);
            }
            if (ComparableDateValue.TryParse(AnnotateAsDateLiteral(parseValue), out var dvComparable))
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
            parseValue = CoerceDateTokenToDouble(parseValue);

            return decimal.TryParse(parseValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var decValue) 
                ? decValue.ToString(CultureInfo.InvariantCulture)
                : double.TryParse(parseValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var value) 
                    ? value.ToString(CultureInfo.InvariantCulture) 
                    : sourceText;
        }

        private static string StringToBoolean(string sourceText)
        {
            var parseValue = RemoveDoubleQuotes(sourceText);
            if (parseValue.Equals(Tokens.True, StringComparison.OrdinalIgnoreCase))
            {
                return Tokens.True;
            }

            if (parseValue.Equals(Tokens.False, StringComparison.OrdinalIgnoreCase))
            {
                return Tokens.False;
            }

            if (parseValue.Equals($"#{Tokens.True}#", StringComparison.Ordinal))
            {
                return Tokens.True;
            }

            if (parseValue.Equals($"#{Tokens.False}#", StringComparison.Ordinal))
            {
                return Tokens.False;
            }

            parseValue = CoerceDateTokenToDouble(parseValue);
            if (double.TryParse(parseValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var asDouble))
            {
                return asDouble != 0 ? Tokens.True : Tokens.False;
            }
            return string.Empty;
        }

        private static string NumericToDate(string source)
        {
            if (double.TryParse(source, NumberStyles.Any, CultureInfo.InvariantCulture, out double dateAsDouble))
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

        private static string CoerceDateTokenToDouble(string sourceText)
        {
            if (sourceText.StartsWith("#") && sourceText.EndsWith("#"))
            {
                if (TryCoerce(sourceText, out ComparableDateValue dvComparable))
                {
                    return Convert.ToDouble(dvComparable.AsDecimal).ToString(CultureInfo.InvariantCulture);
                }
            }
            return sourceText;
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
             => Math.Round(double.Parse(source, CultureInfo.InvariantCulture), MidpointRounding.ToEven).ToString(CultureInfo.InvariantCulture);

        private static string NumericToBoolean(string source)
             => double.Parse(source, CultureInfo.InvariantCulture) != 0 ? Tokens.True : Tokens.False;

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
