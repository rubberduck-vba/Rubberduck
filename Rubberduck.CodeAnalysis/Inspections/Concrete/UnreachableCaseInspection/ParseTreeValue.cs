using Antlr4.Runtime;
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

        public static IParseTreeValue CreateConstant(string value, string declaredType)
        {
            var ptValue = new ParseTreeValue()
            {
                TypeName = declaredType,
                ValueText = value,
                ParsesToConstantValue = true,
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

        private ParseTreeValue()
        {

        }

        public ParseTreeValue(string value, string declaredType)
        {
#if DEBUG
            _inputValue = value;
            _declaredType = declaredType;
#endif
            _valueText = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor");
            TypeName = declaredType ?? throw new ArgumentNullException("null 'declaredType' argument passed to ParseTreeValue constructor");

            _hashCode = value.GetHashCode();
            _valueText = TokenTypeResolver.ConformTokenToType(_valueText, TypeName, out bool parsesToConstant);
            ParsesToConstantValue = parsesToConstant;
            ExceedsTypeRange = ParsesToConstantValue ? ExceedsTypeExtents(_valueText, TypeName) : false;
        }

//        public ParseTreeValue(string value)
//        {
//#if DEBUG
//            _inputValue = value;
//            _declaredType = null;
//#endif
//            _valueText = value ?? throw new ArgumentNullException("null 'value' argument passed to ParseTreeValue constructor");
//            _hashCode = value.GetHashCode();

//            if (TokenTypeResolver.TryDeriveTypeName(value, out string derivedType, out bool derivedFromTypeHint))
//            {
//                TypeName = derivedType;
//                _valueText = derivedFromTypeHint ? RemoveTypeHintChar(value) : value;
//                _valueText = TokenTypeResolver.ConformTokenToType(_valueText, TypeName, out bool parsesToConstant);
//                ParsesToConstantValue = parsesToConstant;
//            }
//        }

        public string TypeName { get; set; } = string.Empty;

        public string ValueText
        {
            private set
            {
                _valueText = value;
            }
            get
            {
                if (ParsesToConstantValue && TypeName != null && TypeName.Equals(Tokens.String))
                {
                    return AnnotateAsStringConstant(_valueText);
                }
                if (ParsesToConstantValue && TypeName != null && TypeName.Equals(Tokens.Date))
                {
                    return AnnotateAsDateLiteral(_valueText);
                }
                return _valueText;
            }
        }

        public bool ParsesToConstantValue { set; get; }

        public bool ExceedsTypeRange { get; set; }

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

        private static string AnnotateAsStringConstant(string input)
        {
            var result = input;
            if (!input.StartsWith("\""))
            {
                result = $"\"{result}";
            }
            if (!input.EndsWith("\""))
            {
                result = $"{result}\"";
            }
            return result;
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
            }
            return false;
        }
    }

    public class ComparableDateValue : IValue, IComparable<ComparableDateValue>
    {
        private readonly DateValue _inner;
        private readonly int _hashCode;

        public ComparableDateValue(DateValue dateValue)
        {
            _inner = dateValue;
            _hashCode = dateValue.AsDecimal.GetHashCode();
        }

        public Parsing.PreProcessing.ValueType ValueType => _inner.ValueType;

        public bool AsBool => _inner.AsBool;

        public byte AsByte => _inner.AsByte;

        public decimal AsDecimal => _inner.AsDecimal;

        public DateTime AsDate => _inner.AsDate;

        public string AsString => _inner.AsString;

        public IEnumerable<IToken> AsTokens => _inner.AsTokens;

        public int CompareTo(ComparableDateValue dateValue)
            => _inner.AsDecimal.CompareTo(dateValue._inner.AsDecimal);

        public override int GetHashCode() => _hashCode;

        public override bool Equals(object obj)
        {
            if (obj is ComparableDateValue decorator)
            {
                return decorator.CompareTo(this) == 0;
            }

            if (obj is DateValue dateValue)
            {
                return dateValue.AsDecimal == _inner.AsDecimal;
            }

            return false;
        }

        public override string ToString()
        {
            return _inner.ToString();
        }
    }

    public static class ParseTreeValueExtensions
    {
        public static bool TryLetCoerce(this IParseTreeValue parseTreeValue, string destinationType, out IParseTreeValue newValue)
        {
            newValue = null;
            var coerce = new LetCoercer(parseTreeValue.TypeName, destinationType);
            if( coerce.TryLetCoerce(parseTreeValue.ValueText, out string valueText))
            {
                newValue = ParseTreeValue.CreateConstant(valueText, destinationType);
                return true;
            }
            return false;
        }

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out long value)
        {
            value = default;
            if (parseTreeValue.TypeName != null && parseTreeValue.TypeName.Equals(Tokens.Date))
            {
                if (parseTreeValue.TryConvertValue(out decimal decValue))
                {
                    return StringValueConverter.TryConvertString(decValue.ToString(), out value, Tokens.Currency);
                }
                return false;
            }
            return StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);
        }

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out double value)
        {
            value = default;
            if (parseTreeValue.TypeName != null && parseTreeValue.TypeName.Equals(Tokens.Date))
            {
                if (parseTreeValue.TryConvertValue(out decimal decValue))
                {
                    return StringValueConverter.TryConvertString(decValue.ToString(), out value);
                }
                return false;
            }
            return StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value, Tokens.Double);
        }

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out decimal value)
        {
            value = default;
            if (parseTreeValue.TypeName != null && parseTreeValue.TypeName.Equals(Tokens.Date))
            {
                if (TryConvertValue(parseTreeValue, out ComparableDateValue dvComparable))
                {
                    value = dvComparable.AsDecimal;
                    return true;
                }
                return false;
            }

            return StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);
        }

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out bool value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out string value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);

        private static bool TryConvertValue(this IParseTreeValue parseTreeValue, out ComparableDateValue value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);
    }
}
