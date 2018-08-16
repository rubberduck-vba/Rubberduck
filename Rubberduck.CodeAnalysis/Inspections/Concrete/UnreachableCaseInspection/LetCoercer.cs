using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public struct LetCoercer
    {
        //Content: Dictionary<sourceTypeName,Dictionary<LetDestinationTypeName,CoercionFunc>
        private static Dictionary<string, Dictionary<string, Func<string, (bool, string)>>> _coercions;

        public static bool TryCoerce((string Type, string Text) source, string destinationType, out string result)
        {
            InitializeCoercions();

            result = string.Empty;
            if (!_coercions.ContainsKey(source.Type))
            {
#if DEBUG
                throw new ArgumentException($"Let Coercion source type: {source.Type} not supported");
#else
                return false;
#endif
            }
            if (!_coercions[source.Type].ContainsKey(destinationType))
            {
#if DEBUG
                throw new ArgumentException($"Let Coercion source=>destination pair: {source.Type}=>{destinationType} not supported");
#else
                return false;
#endif
            }

            var coercer = _coercions[source.Type][destinationType];
            (bool CoercionSuccess, string CoercedValue) = coercer(source.Text);
            result = CoercedValue;
            return CoercionSuccess;
        }

        private static void InitializeCoercions()
        {
            if (_coercions != null) { return; }

            _coercions = new Dictionary<string, Dictionary<string, Func<string, (bool, string)>>>
            {
                [Tokens.String] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { a = RemoveDoubleQuotes(a); return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = (a) => { a = RemoveDoubleQuotes(a); return (Int16.TryParse(a, out _), a); },
                    [Tokens.Long] = (a) => { a = RemoveDoubleQuotes(a); return (Int32.TryParse(a, out _), a); },
                    [Tokens.LongLong] = (a) => { a = RemoveDoubleQuotes(a); return (Int64.TryParse(a, out _), a); },
                    [Tokens.Double] = (a) => { a = RemoveDoubleQuotes(a); return (double.TryParse(a, out _), a); },
                    [Tokens.Single] = (a) => { a = RemoveDoubleQuotes(a); return (float.TryParse(a, out _), a); },
                    [Tokens.Currency] = (a) => { a = RemoveDoubleQuotes(a); return (decimal.TryParse(a, out _), a); },
                    [Tokens.Boolean] = StringToBoolean,
                    [Tokens.Date] = StringToDate,
                },

                [Tokens.Byte] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = Copy,
                    [Tokens.Integer] = Copy,
                    [Tokens.Long] = Copy,
                    [Tokens.LongLong] = Copy,
                    [Tokens.Double] = Copy,
                    [Tokens.Single] = Copy,
                    [Tokens.Currency] = Copy,
                    [Tokens.Boolean] = NumericToBoolean,
                    [Tokens.Date] = NumericToDate,
                },

                [Tokens.Integer] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = Copy,
                    [Tokens.Long] = Copy,
                    [Tokens.LongLong] = Copy,
                    [Tokens.Double] = Copy,
                    [Tokens.Single] = Copy,
                    [Tokens.Currency] = Copy,
                    [Tokens.Boolean] = NumericToBoolean,
                    [Tokens.Date] = NumericToDate,
                },

                [Tokens.Long] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = (a) => { return (Int16.TryParse(a, out _), a); },
                    [Tokens.Long] = Copy,
                    [Tokens.LongLong] = Copy,
                    [Tokens.Double] = Copy,
                    [Tokens.Single] = Copy,
                    [Tokens.Currency] = Copy,
                    [Tokens.Boolean] = NumericToBoolean,
                    [Tokens.Date] = NumericToDate,
                },

                [Tokens.Double] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { a = BankersRound(a); return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = (a) => { a = BankersRound(a); return (Int16.TryParse(a, out _), a); },
                    [Tokens.Long] = (a) => { a = BankersRound(a); return (Int32.TryParse(a, out _), a); },
                    [Tokens.LongLong] = (a) => { a = BankersRound(a); return (long.TryParse(a, out _), a); },
                    [Tokens.Double] = Copy,
                    [Tokens.Single] = (a) => { return (float.TryParse(a, out _), a); },
                    [Tokens.Currency] = (a) => { return (decimal.TryParse(a, out _), a); },
                    [Tokens.Boolean] = NumericToBoolean,
                    [Tokens.Date] = NumericToDate,
                },

                [Tokens.Single] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { a = BankersRound(a); return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = (a) => { a = BankersRound(a); return (Int16.TryParse(a, out _), a); },
                    [Tokens.Long] = (a) => { a = BankersRound(a); return (Int32.TryParse(a, out _), a); },
                    [Tokens.LongLong] = (a) => { a = BankersRound(a); return (long.TryParse(a, out _), a); },
                    [Tokens.Double] = Copy,
                    [Tokens.Single] = Copy,
                    [Tokens.Currency] = (a) => { return (decimal.TryParse(a, out _), a); },
                    [Tokens.Boolean] = NumericToBoolean,
                    [Tokens.Date] = NumericToDate,
                },

                [Tokens.Currency] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { a = BankersRound(a); return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = (a) => { a = BankersRound(a); return (Int16.TryParse(a, out _), a); },
                    [Tokens.Long] = (a) => { a = BankersRound(a); return (Int32.TryParse(a, out _), a); },
                    [Tokens.LongLong] = (a) => { a = BankersRound(a); return (long.TryParse(a, out _), a); },
                    [Tokens.Double] = Copy,
                    [Tokens.Single] = (a) => { return (float.TryParse(a, out _), a); },
                    [Tokens.Currency] = Copy,
                    [Tokens.Boolean] = NumericToBoolean,
                    [Tokens.Date] = NumericToDate,
                },

                [Tokens.Boolean] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = (a) => { var result = BooleanToString(a); return result; },
                    [Tokens.Byte] = (a) => { var val = BooleanAsLong(a); return (true, val != 0 ? byte.MaxValue.ToString() : byte.MinValue.ToString()); },
                    [Tokens.Integer] = (a) => { return (true, BooleanAsLong(a).ToString()); },
                    [Tokens.Long] = (a) => { return (true, BooleanAsLong(a).ToString()); },
                    [Tokens.LongLong] = (a) => { return (true, BooleanAsLong(a).ToString()); },
                    [Tokens.Double] = (a) => { return (true, BooleanAsLong(a).ToString()); },
                    [Tokens.Single] = (a) => { return (true, BooleanAsLong(a).ToString()); },
                    [Tokens.Currency] = (a) => { return (true, BooleanAsLong(a).ToString()); },
                    [Tokens.Boolean] = Copy,
                    [Tokens.Date] = (a) => { var val = BooleanAsLong(a); return NumericToDate(val.ToString()); },
                },

                [Tokens.Date] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = (a) => { return DateToString(a); },
                    [Tokens.Byte] = (a) => { var result = DateToDouble(a); return (byte.TryParse(result.Item2, out _), result.Item2); },
                    [Tokens.Integer] = (a) => { var result = DateToDouble(a); return (Int16.TryParse(result.Item2, out _), result.Item2); },
                    [Tokens.Long] = (a) => { var result = DateToDouble(a); return (Int32.TryParse(result.Item2, out _), result.Item2); },
                    [Tokens.LongLong] = (a) => { var result = DateToDouble(a); return (long.TryParse(result.Item2, out _), result.Item2); },
                    [Tokens.Double] = DateToDouble,
                    [Tokens.Single] = (a) => { var result = DateToDouble(a); return (float.TryParse(result.Item2, out _), result.Item2); },
                    [Tokens.Currency] = (a) => { var result = DateToDouble(a); return (decimal.TryParse(result.Item2, out _), result.Item2); },
                    [Tokens.Boolean] = (a) => { var result = DateToDouble(a); var dbl = double.Parse(result.Item2); return (dbl != 0, dbl != 0 ? Tokens.True : Tokens.False); },
                    [Tokens.Date] = Copy,
                },
            };
        }

        private static (bool, string) StringToDate(string sourceText)
        {
            if (TokenParser.TryParse(AnnotateAsDateLiteral(sourceText), out ComparableDateValue dvComparable))
            {
                return (true, dvComparable.AsDateLiteral());
            }
            if (TokenParser.TryParse(sourceText, out double doubleValue))
            {
                return NumericToDate(sourceText);
            }
            return (false, string.Empty);
        }

        private static (bool, string) StringToBoolean(string sourceText)
        {
            if (sourceText.Equals(Tokens.True, StringComparison.OrdinalIgnoreCase))
            {
                return (true, Tokens.True);
            }
            if (sourceText.Equals(Tokens.False, StringComparison.OrdinalIgnoreCase))
            {
                return (true, Tokens.False);
            }
            if (sourceText.Equals($"#{Tokens.True}#", StringComparison.Ordinal))
            {
                return (true, Tokens.True);
            }
            if (sourceText.Equals($"#{Tokens.False}#", StringComparison.Ordinal))
            {
                return (true, Tokens.False);
            }
            if (double.TryParse(sourceText, out double asDouble))
            {
                return asDouble != 0 ? (true, Tokens.True) : (true, Tokens.False);
            }
            return (false, string.Empty);
        }

        private static (bool, string) NumericToDate(string source)
        {
            if (TokenParser.TryParse(source, out double dateAsDouble))
            {
                var dv = new DateValue(DateTime.FromOADate(dateAsDouble));
                var dateValue = new ComparableDateValue(dv);
                return (true, dateValue.AsDateLiteral());
            }
            return (false, string.Empty);
        }

        private static (bool, string) DateToString(string source)
        {
            return (true, RemoveStartAndEnd(source, "#"));
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

        private static (bool, string) DateToDouble(string source)
        {
            if (TokenParser.TryParse(source, out ComparableDateValue dv))
            {
                return (true, dv.AsDecimal.ToString());
            }
            return (false, string.Empty);
        }

        private static (bool, string) BooleanToString(string source)
        {
            if (source.Equals(Tokens.True) || source.Equals(Tokens.False))
            {
                return Copy(source);
            }

            return (true, double.Parse(source) != 0 ? Tokens.True : Tokens.False);
        }

        private static string BankersRound(string source)
        {
            var parseable = RemoveDoubleQuotes(source);
            if (double.TryParse(source, out double value))
            {
                var integral = Math.Round(value, MidpointRounding.ToEven);
                return integral.ToString();
            }
            throw new OverflowException();
        }

        private static (bool, string) NumericToBoolean(string source)
             => (true, double.Parse(source) != 0 ? Tokens.True : Tokens.False);

        private static long BooleanAsLong(string source)
        {
            if (source.Equals(Tokens.True))
            {
                return -1;
            }
            if (source.Equals(Tokens.False))
            {
                return 0;
            }
            return long.Parse(source);
        }

        private static (bool, string) Copy(string source) => (true, source);

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
