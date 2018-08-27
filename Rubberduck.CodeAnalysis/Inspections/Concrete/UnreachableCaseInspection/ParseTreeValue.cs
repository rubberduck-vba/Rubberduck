using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;
using System.Data;

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
        private readonly int _hashCode;
        private string _valueText;
        private ComparableDateValue _dateValue;
        private StringLiteralExpression _stringConstant;
        private bool? _exceedsTypeRange;

        public static IParseTreeValue CreateValueType(string value, string declaredType)
        {
            if (declaredType == Tokens.Date || declaredType == Tokens.String)
            {
                return new ParseTreeValue(value, declaredType);
            }

            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor"),
                Token = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor"),
                ParsesToConstantValue = true,
                IsOverflowExpression = ExceedsTypeExtents(value, declaredType),
            };
            return ptValue;
        }

        public static IParseTreeValue CreateExpression(string value, string declaredType)
        {
            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor"),
                Token = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor"),
                ParsesToConstantValue = false,
            };
            return ptValue;
        }

        public static IParseTreeValue CreateMismatchExpression(string value, string declaredType)
        {
            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor"),
                Token = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor"),
                ParsesToConstantValue = false,
                IsMismatchExpression = true
            };
            return ptValue;
        }

        public static IParseTreeValue CreateOverflowExpression(string value, string declaredType)
        {
            var ptValue = new ParseTreeValue()
            {
                ValueType = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor"),
                Token = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor"),
                ParsesToConstantValue = false,
                _exceedsTypeRange = true
            };
            return ptValue;
        }

        public ParseTreeValue(string value, string declaredType)
        {
            _valueText = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor");
            ValueType = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor");
            ParsesToConstantValue = false;
            _exceedsTypeRange = null;
            _hashCode = value.GetHashCode();
            _dateValue = null;
            _stringConstant = null;
            IsMismatchExpression = false;

            if (declaredType == Tokens.Date)
            {
                if (LetCoercer.TryCoerce(_valueText, out _dateValue))
                {
                    ParsesToConstantValue = true;
                }
                else
                {
                    throw new ArgumentException($"Unable to coerce {_valueText} to Date");
                }
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
                _exceedsTypeRange = value;
            }
            get
            {
                if (!_exceedsTypeRange.HasValue)
                {
                    _exceedsTypeRange = ParsesToConstantValue ? ExceedsTypeExtents(_valueText, ValueType) : false;
                }
                return _exceedsTypeRange.Value;
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

        private static Dictionary<string, Action<string>> OverflowChecks = new Dictionary<string, Action<string>>()
        {
            [Tokens.Byte] = (a) => { byte.Parse(a); },
            [Tokens.Integer] = (a) => { Int16.Parse(a); },
            [Tokens.Long] = (a) => { Int32.Parse(a); },
            [Tokens.LongLong] = (a) => { Int64.Parse(a); },
            [Tokens.Double] = (a) => { double.Parse(a); },
            [Tokens.Single] = (a) => { float.Parse(a); },
            [Tokens.Currency] = (a) => { var value = decimal.Parse(a); if (value < VBACurrency.MinValue || value > VBACurrency.MaxValue) { throw new OverflowException(); } },
            [Tokens.Boolean] = (a) => { if (!(a.Equals(Tokens.True) || a.Equals(Tokens.False))) { long.Parse(a); } },
        };

        private static bool ExceedsTypeExtents(string valueText, string typeName)
        {
            if (OverflowChecks.ContainsKey(typeName))
            {
                try
                {
                    OverflowChecks[typeName](valueText);
                }
                catch (OverflowException)
                {
                    return true;
                }
                catch (FormatException)
                {
                    return false;
                }
            }
            return false;
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
