using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValue
    {
        string Token { get; }
        string ValueType { get; }
        bool ParsesToConstantValue { get; }
        bool IsOverflowExpression { get; }
        bool IsMismatchExpression { get; }
    }

    public struct ParseTreeValue : IParseTreeValue
    {
        private const string NULL_VALUETYPE_MSG = "null 'valueType' argument passed to ParseTreeValue constructor";
        private const string NULL_TOKEN_MSG = "null 'token' argument passed to ParseTreeValue constructor";

        private readonly int _hashCode;
        private string _valueText;
        private ComparableDateValue _dateValue;
        private StringLiteralExpression _stringConstant;
        private bool? _exceedsValueTypeRange;

        public static IParseTreeValue CreateValueType(string token, string declaredValueType)
        {
            if (declaredValueType == Tokens.Date || declaredValueType == Tokens.String)
            {
                return new ParseTreeValue(token, declaredValueType);
            }

            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredValueType ?? throw new ArgumentNullException(NULL_VALUETYPE_MSG),
                Token = token ?? throw new ArgumentNullException(NULL_TOKEN_MSG),
                ParsesToConstantValue = true,
                IsOverflowExpression = LetCoercer.ExceedsValueTypeRange(declaredValueType, token),
            };
            return ptValue;
        }

        public static IParseTreeValue CreateExpression(string value, string declaredValueType)
        {
            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredValueType ?? throw new ArgumentNullException(NULL_VALUETYPE_MSG),
                Token = value ?? throw new ArgumentNullException(NULL_TOKEN_MSG),
                ParsesToConstantValue = false,
            };
            return ptValue;
        }

        public static IParseTreeValue CreateMismatchExpression(string value, string declaredValueType)
        {
            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredValueType ?? throw new ArgumentNullException(NULL_VALUETYPE_MSG),
                Token = value ?? throw new ArgumentNullException(NULL_TOKEN_MSG),
                ParsesToConstantValue = false,
                IsMismatchExpression = true
            };
            return ptValue;
        }

        public static IParseTreeValue CreateOverflowExpression(string value, string declaredValueType)
        {
            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredValueType ?? throw new ArgumentNullException(NULL_VALUETYPE_MSG),
                Token = value ?? throw new ArgumentNullException(NULL_TOKEN_MSG),
                ParsesToConstantValue = false,
                _exceedsValueTypeRange = true
            };
            return ptValue;
        }

        public ParseTreeValue(string value, string declaredType)
        {
            ValueType = declaredType ?? throw new ArgumentNullException(NULL_VALUETYPE_MSG);
            _valueText = value ?? throw new ArgumentNullException(NULL_TOKEN_MSG);
            ParsesToConstantValue = false;
            _exceedsValueTypeRange = null;
            _hashCode = value.GetHashCode();
            _dateValue = null;
            _stringConstant = null;
            IsMismatchExpression = false;

            if (declaredType == Tokens.Date)
            {
                ParsesToConstantValue = LetCoercer.TryCoerce(_valueText, out _dateValue);
            }
            if (declaredType.Equals(Tokens.String))
            {
                if (_valueText.StartsWith("\"") && _valueText.EndsWith("\""))
                {
                    _stringConstant = new StringLiteralExpression(new ConstantExpression(new StringValue(_valueText)));
                    ParsesToConstantValue = true;
                }
            }
        }

        public string ValueType { private set;  get; }

        public string Token
        {
            private set
            {
                _valueText = value;
            }
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
                return _valueText;
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
                    _exceedsValueTypeRange = ParsesToConstantValue ? LetCoercer.ExceedsValueTypeRange(ValueType, _valueText) : false;
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
    }

    public static class ParseTreeValueExtensions
    {
        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, string destinationType, out IParseTreeValue newValue)
        {
            newValue = null;
            if (LetCoercer.TryCoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), destinationType, out string valueText))
            {
                newValue = ParseTreeValue.CreateValueType(valueText, destinationType);
                return true;
            }
            return false;
        }

        public static double AsDouble(this IParseTreeValue parseTreeValue)
            => double.Parse(LetCoercer.CoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Double));

        public static decimal AsCurrency(this IParseTreeValue parseTreeValue)
            => decimal.Parse(LetCoercer.CoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Currency));

        public static long AsLong(this IParseTreeValue parseTreeValue)
            => long.Parse(LetCoercer.CoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Long));

        public static bool AsBoolean(this IParseTreeValue parseTreeValue)
            =>LetCoercer.CoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Boolean).Equals(Tokens.True);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out long newValue)
            => TryLetCoerce(parseTreeValue, long.Parse, Tokens.Long, out newValue);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out double newValue)
        => TryLetCoerce(parseTreeValue, double.Parse, Tokens.Double, out newValue);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out decimal newValue)
            => TryLetCoerce(parseTreeValue, decimal.Parse, Tokens.Currency, out newValue);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out bool value)
            => TryLetCoerce(parseTreeValue, bool.Parse, Tokens.Boolean, out value);

        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out string value)
            => TryLetCoerce(parseTreeValue, (a) => { return a; }, Tokens.String, out value);

        private static bool TryLetCoerce(this IParseTreeValue parseTreeValue, out ComparableDateValue value)
            => TryLetCoerceToDate(parseTreeValue, out value);

        private static bool TryLetCoerce<T>(this IParseTreeValue parseTreeValue, Func<string, T> parser, string typeName, out T newValue)
        {
            newValue = default;
            if (LetCoercer.TryCoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), typeName, out string valueText))
            {
                newValue = parser(valueText);
                return true;
            }
            return false;
        }

        private static bool TryLetCoerceToDate(IParseTreeValue parseTreeValue, out ComparableDateValue value)
        {
            value = default;
            if (LetCoercer.TryCoerceToken((parseTreeValue.ValueType, parseTreeValue.Token), Tokens.Date, out string valueText))
            {
                var literal = new DateLiteralExpression(new ConstantExpression(new StringValue(valueText)));
                value = new ComparableDateValue((DateValue)literal.Evaluate());
                return true;
            }
            return false;
        }
    }
}
