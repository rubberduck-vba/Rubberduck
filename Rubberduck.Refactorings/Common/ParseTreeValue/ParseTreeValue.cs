using System;
using System.Collections.Generic;
using System.Globalization;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Refactorings;

namespace Rubberduck.Refactoring.ParseTreeValue
{
    public struct ParseTreeValue : IParseTreeValue
    {
        private readonly int _hashCode;
        private TypeTokenPair _typeTokenPair;
        private ComparableDateValue _dateValue;
        private StringLiteralExpression _stringConstant;
        private bool? _exceedsValueTypeRange;

        private static Dictionary<string,string> ControlCharacterCompareTokens = new Dictionary<string, string>()
        {
            ["Chr$(8)"] = "Chr$(8)", //vbBack
            ["Chr$(13)"] = "Chr$(13)", //vbCr
            ["Chr$(13) + Chr$(10)"] = "Chr$(13)Chr$(10)", //vbCrLf
            ["Chr$(10)"] = "Chr$(10)", //vbLf
            ["Chr$(12)"] = "Chr$(12)", //vbFormFeed
            ["Chr$(13) & Chr$(10)"] = "Chr$(13)Chr$(10)", //vbNewLine
            ["Chr$(0)"] = "Chr$(0)", //vbNullChar
            ["Chr$(9)"] = "Chr$(9)", //vbTab
            ["Chr$(11)"] = "Chr$(11)", //vbVerticalTab
            ["Chr$(13)Chr$(10)"] = "Chr$(13)Chr$(10)",
        };

        public static bool TryGetNonPrintingControlCharCompareToken(string controlCharCandidate, out string comparableToken)
        {
            comparableToken = controlCharCandidate;
            if (controlCharCandidate.StartsWith(Tokens.Chr))
            {
                var key = controlCharCandidate.Replace("Chr(", "Chr$(");
                return ControlCharacterCompareTokens.TryGetValue(key, out comparableToken);
            }
            return false;
        }

        public static IParseTreeValue CreateValueType(TypeTokenPair value)
        {
            if (value.ValueType.Equals(Tokens.Date) || value.ValueType.Equals(Tokens.String))
            {
                return new ParseTreeValue(value);
            }

            var ptValue = new ParseTreeValue()
            {
                _typeTokenPair = value,
                ParsesToConstantValue = true,
                IsOverflowExpression = LetCoerce.ExceedsValueTypeRange(value.ValueType, value.Token),
            };
            return ptValue;
        }

        public static IParseTreeValue CreateExpression(TypeTokenPair typeToken)
        {
            var ptValue = new ParseTreeValue()
            {
                _typeTokenPair = typeToken,
                ParsesToConstantValue = false,
            };
            return ptValue;
        }

        public static IParseTreeValue CreateMismatchExpression(string value, string declaredValueType)
        {
            var ptValue = new ParseTreeValue()
            {
                _typeTokenPair = new TypeTokenPair(declaredValueType, value),
                ParsesToConstantValue = false,
                IsMismatchExpression = true
            };
            return ptValue;
        }

        public static IParseTreeValue CreateOverflowExpression(string value, string declaredValueType)
        {
            var ptValue = new ParseTreeValue()
            {
                _typeTokenPair = new TypeTokenPair(declaredValueType, value),
                ParsesToConstantValue = false,
                _exceedsValueTypeRange = true
            };
            return ptValue;
        }

        public ParseTreeValue(TypeTokenPair valuePair)
        {
            _typeTokenPair = valuePair;
            ParsesToConstantValue = false;
            _exceedsValueTypeRange = null;
            _hashCode = valuePair.Token.GetHashCode();
            _dateValue = null;
            _stringConstant = null;
            IsMismatchExpression = false;

            if (valuePair.ValueType.Equals(Tokens.Date))
            {
                ParsesToConstantValue = LetCoerce.TryCoerce(_typeTokenPair.Token, out _dateValue);
            }
            else if (valuePair.ValueType.Equals(Tokens.String) && IsStringConstant(valuePair.Token))
            {
                _stringConstant = new StringLiteralExpression(new ConstantExpression(new StringValue(_typeTokenPair.Token)));
                ParsesToConstantValue = true;
            }
            else if (valuePair.ValueType.Equals(Tokens.String)
                &&  TryGetNonPrintingControlCharCompareToken(valuePair.Token, out _))
            {
                ParsesToConstantValue = true;
            }
        }

        public string ValueType => _typeTokenPair.ValueType;

        public string Token
        {
            get
            {
                if (_dateValue != null)
                {
                    return _dateValue.AsDateLiteral();
                }
                if (_stringConstant != null)
                {
                    return $"\"{_stringConstant.Evaluate().AsString}\"";
                }
                return _typeTokenPair.Token;
            }
        }

        public bool ParsesToConstantValue { private set; get; }

        public bool IsMismatchExpression { private set; get; }

        public bool IsOverflowExpression
        {
            private set
            {
                _exceedsValueTypeRange = value;
            }
            get
            {
                if (!_exceedsValueTypeRange.HasValue)
                {
                    _exceedsValueTypeRange = ParsesToConstantValue ? LetCoerce.ExceedsValueTypeRange(ValueType, _typeTokenPair.Token) : false;
                }
                return _exceedsValueTypeRange.Value;
            }
        }

        public override string ToString() => Token;

        public override bool Equals(object obj)
        {
            if (obj is ParseTreeValue ptValue)
            {
                return ptValue.Token == Token && ptValue.ValueType == ValueType;
            }

            return false;
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        private static bool IsStringConstant(string candidate) => candidate.StartsWith("\"") && candidate.EndsWith("\"");
    }

    public static class ParseTreeValueExtensions
    {
        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, string destinationType, out IParseTreeValue newValue)
        {
            newValue = null;
            if (LetCoerce.TryCoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), destinationType, out var valueText))
            {
                newValue = ParseTreeValue.CreateValueType(new TypeTokenPair(destinationType, valueText));
                return true;
            }
            return false;
        }

        public static double AsDouble(this IParseTreeValue parseTreeValue)
            => double.Parse(LetCoerce.Coerce((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Double), CultureInfo.InvariantCulture);

        public static decimal AsCurrency(this IParseTreeValue parseTreeValue)
            => decimal.Parse(LetCoerce.Coerce((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Currency), CultureInfo.InvariantCulture);

        public static long AsLong(this IParseTreeValue parseTreeValue)
            => long.Parse(LetCoerce.Coerce((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Long), CultureInfo.InvariantCulture);

        public static bool AsBoolean(this IParseTreeValue parseTreeValue)
            => LetCoerce.Coerce((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Boolean).Equals(Tokens.True);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out long newValue)
            => TryLetCoerce(parseTreeValue, s => long.Parse(s, CultureInfo.InvariantCulture), Tokens.Long, out newValue);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out double newValue)
        => TryLetCoerce(parseTreeValue, s => double.Parse(s, CultureInfo.InvariantCulture), Tokens.Double, out newValue);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out decimal newValue)
            => TryLetCoerce(parseTreeValue, s => decimal.Parse(s, CultureInfo.InvariantCulture), Tokens.Currency, out newValue);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out bool value)
            => TryLetCoerce(parseTreeValue, bool.Parse, Tokens.Boolean, out value);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out string value)
            => TryLetCoerce(parseTreeValue, a => a, Tokens.String, out value);

        private static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out ComparableDateValue value)
            => TryLetCoerceToDate(parseTreeValue, out value);

        private static bool TryLetCoerce<T>(this IParseTreeValue parseTreeValue, Func<string, T> parser, string typeName, out T newValue)
        {
            newValue = default;
            if (LetCoerce.TryCoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), typeName, out string valueText))
            {
                newValue = parser(valueText);
                return true;
            }
            return false;
        }

        private static bool TryLetCoerceToDate(IParseTreeValue parseTreeValue, out ComparableDateValue value)
        {
            value = default;
            if (LetCoerce.TryCoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Date, out string valueText))
            {
                var literal = new DateLiteralExpression(new ConstantExpression(new StringValue(valueText)));
                value = new ComparableDateValue((DateValue)literal.Evaluate());
                return true;
            }
            return false;
        }
    }
}
