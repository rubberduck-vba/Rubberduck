using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public struct LetCoercer
    {
        private readonly string _sourceType;
        private readonly string _destinationType;
        private static Dictionary<string, Dictionary<string, Func<string, (bool, string)>>> _coercions; 

        public LetCoercer(string sourceType, string destinationType)
        {
            _sourceType = sourceType;
            _destinationType = destinationType;
            InitializeCoercions();
        }

        public bool TryLetCoerce(string sourceText, out string result)
        {
            result = string.Empty;
            if (!_coercions.ContainsKey(_sourceType))
            {
#if DEBUG
                throw new ArgumentException($"Let Coercion source type: {_sourceType} not supported");
#else
                return false;
#endif
            }
            if (!_coercions[_sourceType].ContainsKey(_destinationType))
            {
#if DEBUG
                throw new ArgumentException($"Let Coercion source=>destination pair: {_sourceType}=>{_destinationType} not supported");
#else
                return false;
#endif
            }
            if (!_coercions.ContainsKey(_sourceType))
            {
                return false;
            }

            var coercer = _coercions[_sourceType][_destinationType];
            var results = coercer(sourceText);
            result = results.Item2;
            return results.Item1;
        }

        private static (bool, string) StringToDate(string sourceText)
        {
            var candidate = AnnotateAsDateLiteral(sourceText);
            if (StringValueConverter.TryConvertString(candidate, out ComparableDateValue dvComparable))
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
            if (sourceText.Equals($"#{Tokens.True}#",StringComparison.Ordinal))
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
            if (StringValueConverter.TryConvertString(source, out double dateAsDouble))
            {
                var dv = new DateValue(DateTime.FromOADate(dateAsDouble));
                var dateValue = new ComparableDateValue(dv);
                return (true, AnnotateAsDateLiteral(dateValue.AsDate.ToString(CultureInfo.InvariantCulture)));
            }
            return (false, string.Empty);
        }

        private static (bool, string) DateToString(string source)
        {
            return (true,RemoveStartAndEnd(source, "#"));
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
            if (StringValueConverter.TryConvertString(source, out ComparableDateValue dv))
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
            return long.Parse(source); //) ? (true, value != 0 ? "-1" : "0") : (false,string.Empty);
        }

        private static (bool,string) Copy(string source) => (true, source);

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

        private void InitializeCoercions()
        {
            if (_coercions is null)
            {
                //Dictionary<sourceTypeName,Dictionary<LetdestinationTypeName,Func>
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
                        [Tokens.String] = (a) => { var result = BooleanToString(a);  return result; },
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
                        [Tokens.String] = (a) => { return  DateToString(a); },
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

    public class TokenTypeResolver
    {
        public static string DeriveTypeName(string inputString, out bool derivedFromTypeHint)
        {
            if (TryDeriveTypeName(inputString, out string result, out derivedFromTypeHint))
            {
                return result;
            }
            return string.Empty;
        }

        public static bool TryDeriveTypeName(string inputString, out string typeName, out bool derivedFromTypeHint)
        {
            derivedFromTypeHint = false;
            typeName = string.Empty;

            if (inputString is null || inputString.Length == 0)
            {
                return false;
            }

            if (TryParseAsDateLiteral(inputString, out ComparableDateValue _))
            {
                typeName = Tokens.Date;
                return true;
            }

            if (SymbolList.TypeHintToTypeName.TryGetValue(inputString.Last().ToString(), out string hintResult))
            {
                derivedFromTypeHint = true;
                typeName = hintResult;
                return true;
            }

            if (IsStringConstant(inputString))
            {
                typeName = Tokens.String;
                return true;
            }

            if (inputString.Contains(".") || inputString.Count(ch => ch.Equals('E')) == 1)
            {
                if (double.TryParse(inputString, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                {
                    typeName = Tokens.Double;
                    return true;
                }
            }

            if (inputString.Equals(Tokens.True) || inputString.Equals(Tokens.False))
            {
                typeName = Tokens.Boolean;
                return true;
            }

            if (short.TryParse(inputString, out _)
                || TryParseAsHexLiteral(inputString, out short _)
                || TryParseAsOctalLiteral(inputString, out short _))
            {
                typeName = Tokens.Integer;
                return true;
            }

            if (int.TryParse(inputString, out _)
                || TryParseAsHexLiteral(inputString, out int _)
                || TryParseAsOctalLiteral(inputString, out int _))
            {
                typeName = Tokens.Long;
                return true;
            }

            if (long.TryParse(inputString, out _))
            {
                typeName = Tokens.Double; //See 3.3.2 of the VBA specification.
                return true;
            }

            return false;
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

            if (/*typeHintType == null ||*/ typeHintType == Tokens.String) // || IsStringConstant(inputString))
            {
                //if (IsStringConstant(inputString))
                {
                    //result = (Tokens.String, inputString);
                    return true;
                }
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

            //if (inputString.Contains(".") || inputString.Count(ch => ch.Equals('E')) == 1)
            //{
            //    if (double.TryParse(inputString, NumberStyles.Float, CultureInfo.InvariantCulture, out double dVal))
            //    {
            //        result = (Tokens.Double, dVal.ToString());
            //        return true;
            //    }
            //}

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

            return derivedFromTypeHint; //false;
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
            var candidate = AnnotateAsDateLiteral(sourceText);
            if (StringValueConverter.TryConvertString(candidate, out ComparableDateValue dvComparable))
            {
                var result = dvComparable.AsDate.ToString(CultureInfo.InvariantCulture);
                result = AnnotateAsDateLiteral(sourceText);
                return (true, result);
            }
            return (false, sourceText);
        }


        public static (bool Success, string CoercedText) LetCoerce((string Type, string Text) source, string destinationType)
        {
            var result = string.Empty;
            var coercer = new LetCoercer(source.Type, destinationType);
            var returnValue = coercer.TryLetCoerce(source.Text, out string coercedValue);
            return (returnValue, coercedValue);
        }

        public static string ConformTokenToType(string valueToken, string conformTypeName, out bool parsesToConstant)
        {
            parsesToConstant = false;

            if (valueToken.Equals(double.NaN.ToString(CultureInfo.InvariantCulture)) &&
                !conformTypeName.Equals(Tokens.String))
            {
                return "";
            }

            else if (conformTypeName.Equals(Tokens.LongLong) || conformTypeName.Equals(Tokens.Long) ||
                conformTypeName.Equals(Tokens.Integer) || conformTypeName.Equals(Tokens.Byte))
            {
                if (StringValueConverter.TryConvertString(valueToken, out long newVal))
                {
                    valueToken = newVal.ToString();
                    parsesToConstant = true;
                }

                if (conformTypeName.Equals(Tokens.Integer))
                {
                    if (TryParseAsHexLiteral(valueToken, out short outputHex))
                    {
                        valueToken = outputHex.ToString();
                        parsesToConstant = true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out short outputOctal))
                    {
                        parsesToConstant = true;
                        valueToken = outputOctal.ToString();
                    }
                }

                if (conformTypeName.Equals(Tokens.Long))
                {
                    if (TryParseAsHexLiteral(valueToken, out int outputHex))
                    {
                        valueToken = outputHex.ToString();
                        parsesToConstant = true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out int outputOctal))
                    {
                        valueToken = outputOctal.ToString();
                        parsesToConstant = true;
                    }
                }

                if (conformTypeName.Equals(Tokens.LongLong))
                {
                    if (TryParseAsHexLiteral(valueToken, out long outputHex))
                    {
                        valueToken = outputHex.ToString();
                        parsesToConstant = true;
                    }

                    if (TryParseAsOctalLiteral(valueToken, out long outputOctal))
                    {
                        valueToken = outputOctal.ToString();
                        parsesToConstant = true;
                    }
                }
            }

            else if (conformTypeName.Equals(Tokens.Double) || conformTypeName.Equals(Tokens.Single))
            {
                var derivedTypeName = DeriveTypeName(valueToken, out bool usedTypehint);
                if (derivedTypeName.Equals(Tokens.Date))
                {
                    if (StringValueConverter.TryConvertString(valueToken, out ComparableDateValue dvComparable))
                    {
                        valueToken = Convert.ToDouble(dvComparable.AsDecimal).ToString(CultureInfo.InvariantCulture);
                        parsesToConstant = true;
                    }
                }
                else if (StringValueConverter.TryConvertString(valueToken, out double newVal))
                {
                    valueToken = newVal.ToString(CultureInfo.InvariantCulture);
                    parsesToConstant = true;
                }
            }

            else if (conformTypeName.Equals(Tokens.Boolean))
            {
                if (StringValueConverter.TryConvertString(valueToken, out bool newVal))
                {
                    valueToken = newVal.ToString();
                    parsesToConstant = true;
                }
            }

            else if (conformTypeName.Equals(Tokens.String))
            {
                parsesToConstant = IsStringConstant(valueToken);
            }

            else if (conformTypeName.Equals(Tokens.Currency))
            {
                if (StringValueConverter.TryConvertString(valueToken, out decimal newVal))
                {
                    var currencyValue = Math.Round(newVal, 4, MidpointRounding.ToEven);
                    valueToken = currencyValue.ToString(CultureInfo.InvariantCulture);
                    parsesToConstant = true;
                }
            }

            else if (conformTypeName.Equals(Tokens.Date))
            {
                var derivedTypeName = DeriveTypeName(valueToken, out bool _);
                if (!derivedTypeName.Equals(Tokens.Date))
                {
                    var sourceType = derivedTypeName.Equals(string.Empty) ? Tokens.String : derivedTypeName;
                    var (Success, CoercedText) = LetCoerce((sourceType, valueToken), Tokens.Date);
                    valueToken = CoercedText;
                    parsesToConstant = Success;
                }
                else if (StringValueConverter.TryConvertString(valueToken, out ComparableDateValue dvComparable))
                {
                    valueToken = dvComparable.AsDate.ToString(CultureInfo.InvariantCulture);
                    valueToken = AnnotateAsDateLiteral(valueToken);
                    parsesToConstant = true;
                }
            }
            return valueToken;
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

        private static bool TryParseAsDateLiteral(string valueString, out ComparableDateValue value)
        {
            return StringValueConverter.TryConvertString(valueString, out value);
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
