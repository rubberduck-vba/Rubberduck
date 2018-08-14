using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public struct LetCoercer
    {
        //Dictionary<sourceTypeName,Dictionary<LetdestinationTypeName,Func>
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
            if (!_coercions.ContainsKey(source.Type))
            {
                return false;
            }

            var coercer = _coercions[source.Type][destinationType];
            var results = coercer(source.Text);
            result = results.Item2;
            return results.Item1;
        }

        private static (bool, string) StringToDate(string sourceText)
        {
            var candidate = AnnotateAsDateLiteral(sourceText);
            if (TokenParser.TryParse(candidate, out ComparableDateValue dvComparable))
            {
                var result = dvComparable.AsDate.ToString(CultureInfo.InvariantCulture);
                result = AnnotateAsDateLiteral(sourceText);
                return (true, result);
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
                return (true, AnnotateAsDateLiteral(dateValue.AsDate.ToString(CultureInfo.InvariantCulture)));
            }
            return (false, string.Empty);
        }

        private static (bool, string) DateToString(string source)
        {
            return (true, RemoveStartAndEnd(source, "#"));
        }

        private static string RemoveQuotes(string source)
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

            var dValue = double.Parse(source);
            return (true, dValue != 0 ? Tokens.True : Tokens.False);
        }

        private static string BankersRound(string source)
        {
            var parseable = RemoveQuotes(source);
            if (double.TryParse(source, out double value))
            {
                var integral = Math.Round(value, MidpointRounding.ToEven);
                return integral.ToString();
            }
            throw new OverflowException();
        }

        private static (bool, string) NumericToBoolean(string source)
        {
            double.TryParse(source, out double value);
            return (value != 0, value != 0 ? Tokens.True : Tokens.False);
        }

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

        private static (bool, string) IntegralToGreater(string source)
         => (long.TryParse(source, out _), source.ToString());

        private static (bool, string) IntegralToRational(string source)
        {
            double.TryParse(source, out double value);
            return (true, value.ToString());
        }

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
            result.Replace(" 00:00:00", "");
            return result;
        }

        private static void InitializeCoercions()
        {
            if (_coercions != null) { return; }

            _coercions = new Dictionary<string, Dictionary<string, Func<string, (bool, string)>>>
            {
                [Tokens.String] = new Dictionary<string, Func<string, (bool, string)>>
                {
                    [Tokens.String] = Copy,
                    [Tokens.Byte] = (a) => { a = RemoveQuotes(a); return (byte.TryParse(a, out _), a); },
                    [Tokens.Integer] = (a) => { a = RemoveQuotes(a); return (Int16.TryParse(a, out _), a); },
                    [Tokens.Long] = (a) => { a = RemoveQuotes(a); return (Int32.TryParse(a, out _), a); },
                    [Tokens.LongLong] = (a) => { a = RemoveQuotes(a); return (Int64.TryParse(a, out _), a); },
                    [Tokens.Double] = (a) => { a = RemoveQuotes(a); return (double.TryParse(a, out _), a); },
                    [Tokens.Single] = (a) => { a = RemoveQuotes(a); return (float.TryParse(a, out _), a); },
                    [Tokens.Currency] = (a) => { a = RemoveQuotes(a); return (decimal.TryParse(a, out _), a); },
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
                    [Tokens.Integer] = (a) => { var val = BooleanAsLong(a); return Copy(val.ToString()); },
                    [Tokens.Long] = (a) => { var val = BooleanAsLong(a); return Copy(val.ToString()); },
                    [Tokens.LongLong] = (a) => { var val = BooleanAsLong(a); return Copy(val.ToString()); },
                    [Tokens.Double] = (a) => { var val = BooleanAsLong(a); return Copy(val.ToString()); },
                    [Tokens.Single] = (a) => { var val = BooleanAsLong(a); return Copy(val.ToString()); },
                    [Tokens.Currency] = (a) => { var val = BooleanAsLong(a); return Copy(val.ToString()); },
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
    }
}
