using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValue
    {
        string ValueText { get; }
        string TypeName { get; }
        bool ParsesToConstantValue { get; set; }
        bool ExceedsTypeRange { get; set; }
    }

    public class ParseTreeValue : IParseTreeValue
    {
        private static decimal CURRENCYMIN = -922337203685477.5808M;
        private static decimal CURRENCYMAX = 922337203685477.5807M;

#if DEBUG  //useful when debugging
        private readonly string _inputValue;
        private readonly string _declaredType;
#endif

        private readonly int _hashCode;

        private string _valueText;
        private ComparableDateValue _dateValue;
        private StringLiteralExpression _stringConstant;
        private bool? _exceedsTypeRange;

        public static IParseTreeValue CreateConstant(string value, string declaredType)
        {
            if (declaredType == Tokens.Date || declaredType == Tokens.String)
            {
                return new ParseTreeValue(value, declaredType);
            }

            var ptValue = new ParseTreeValue()
            {
                TypeName = declaredType,
                ValueText = value,
                ParsesToConstantValue = true,
                ExceedsTypeRange = ExceedsTypeExtents(value, declaredType),
            };
            return ptValue;
        }

        public static IParseTreeValue CreateExpression(string value, string declaredType)
        {
            var ptValue = new ParseTreeValue()
            {
                TypeName = declaredType,
                ValueText = value,
                ParsesToConstantValue = false,
            };
            return ptValue;
        }

        private ParseTreeValue(){ }

        public ParseTreeValue(string value, string declaredType)
        {
#if DEBUG
            _inputValue = value;
            _declaredType = declaredType;
#endif
            _valueText = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor");
            TypeName = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor");

            _hashCode = value.GetHashCode();

            if (declaredType == Tokens.Date)
            {
                InitializeDateValueMember();
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

        private void InitializeDateValueMember()
        {
            if (TokenParser.TryParse(_valueText, out ComparableDateValue dv))
            {
                _dateValue = dv;
                ParsesToConstantValue = true;
            }
            else if (TokenParser.TryParse(_valueText, out double dblVal))
            {
                var dateValue = new DateValue(DateTime.FromOADate(dblVal));
                _dateValue = new ComparableDateValue(dateValue);
                ParsesToConstantValue = true;
            }
            else
            {
                throw new ArgumentException($"Unable to convert {_valueText} to Date");
            }
        }

        public string TypeName { get; set; } = string.Empty;

        public string ValueText
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

        public bool ParsesToConstantValue { get; set; }

        public bool ExceedsTypeRange
        {
            get
            {
                if (!_exceedsTypeRange.HasValue)
                {
                    _exceedsTypeRange = ParsesToConstantValue ? ExceedsTypeExtents(_valueText, TypeName) : false;
                }
                return _exceedsTypeRange.Value;
            }
            set
            {
                _exceedsTypeRange = value;
            }
        }

        public override string ToString() => ValueText;

        public override bool Equals(object obj)
        {
            if (obj is ParseTreeValue ptValue)
            {
                return ptValue.ValueText == ValueText && ptValue.TypeName == TypeName;
            }

            return false;
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        private static string RemoveTypeHintChar(string inputValue)
        {
            if (inputValue == string.Empty)
            {
                return inputValue;
            }

            return SymbolList.TypeHintToTypeName.ContainsKey(inputValue.Last().ToString())
                ? inputValue.Substring(0, inputValue.Length - 1)
                : inputValue;
        }

        private static Dictionary<string, Action<string>> OverflowChecks = new Dictionary<string, Action<string>>()
        {
            [Tokens.Byte] = (a) => { byte.Parse(a); },
            [Tokens.Integer] = (a) => { Int16.Parse(a); },
            [Tokens.Long] = (a) => { Int32.Parse(a); },
            [Tokens.LongLong] = (a) => { Int64.Parse(a); },
            [Tokens.Double] = (a) => { double.Parse(a); },
            [Tokens.Single] = (a) => { float.Parse(a); },
            [Tokens.Currency] = (a) => { var value = decimal.Parse(a); if (value < CURRENCYMIN || value > CURRENCYMAX) { throw new OverflowException(); } },
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
                catch(FormatException)
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
            if(LetCoercer.TryLetCoerce((parseTreeValue.TypeName, parseTreeValue.ValueText), destinationType, out string valueText))
            {
                newValue = ParseTreeValue.CreateConstant(valueText, destinationType);
                return true;
            }
            return false;
        }

        public static bool TryParse(this IParseTreeValue parseTreeValue, out long value)
        {
            value = default;
            if (parseTreeValue.TypeName != null && parseTreeValue.TypeName.Equals(Tokens.Date))
            {
                if (parseTreeValue.TryParse(out decimal decValue))
                {
                    return TokenParser.TryParse(decValue.ToString(), out value, Tokens.Currency);
                }
                return false;
            }
            return TokenParser.TryParse(parseTreeValue.ValueText, out value);
        }

        public static bool TryParse(this IParseTreeValue parseTreeValue, out double value)
        {
            value = default;
            if (parseTreeValue.TypeName != null && parseTreeValue.TypeName.Equals(Tokens.Date))
            {
                if (parseTreeValue.TryParse(out decimal decValue))
                {
                    return TokenParser.TryParse(decValue.ToString(), out value);
                }
                return false;
            }
            return TokenParser.TryParse(parseTreeValue.ValueText, out value, Tokens.Double);
        }

        public static bool TryParse(this IParseTreeValue parseTreeValue, out decimal value)
        {
            value = default;
            if (parseTreeValue.TypeName != null && parseTreeValue.TypeName.Equals(Tokens.Date))
            {
                if (TryParse(parseTreeValue, out ComparableDateValue dvComparable))
                {
                    value = dvComparable.AsDecimal;
                    return true;
                }
                return false;
            }

            return TokenParser.TryParse(parseTreeValue.ValueText, out value);
        }

        public static bool TryParse(this IParseTreeValue parseTreeValue, out bool value)
            => TokenParser.TryParse(parseTreeValue.ValueText, out value);

        public static bool TryParse(this IParseTreeValue parseTreeValue, out string value)
            => TokenParser.TryParse(parseTreeValue.ValueText, out value);

        private static bool TryParse(this IParseTreeValue parseTreeValue, out ComparableDateValue value)
            => TokenParser.TryParse(parseTreeValue.ValueText, out value);
    }
}
