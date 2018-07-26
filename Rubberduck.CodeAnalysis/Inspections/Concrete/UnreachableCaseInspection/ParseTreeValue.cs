using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValue
    {
        string ValueText { get; }
        string TypeName { get; }
        bool ParsesToConstantValue { get; set; }
        decimal AsCurrency { get; }
        bool IsOverflowException { get; set; }
    }

    public class ParseTreeValue : IParseTreeValue
    {
        private static decimal CURRENCYMIN = -922337203685477.5808M;
        private static decimal CURRENCYMAX = 922337203685477.5807M;

        private readonly string _inputValue;
        private readonly string _declaredType;
        private readonly string _derivedType;
        private readonly int _hashCode;

        private string _valueText;
        private ComparableDateValue _dateValue;

        public ParseTreeValue(string value, string declaredType = null)
        {
            if (value is null)
            {
                throw new ArgumentNullException("null 'value' argument passed to UCIValue");
            }

            _inputValue = value;
            _hashCode = value.GetHashCode();
            ValueText = value;
            ParsesToConstantValue = IsStringConstant(value);
            _declaredType = ParsesToConstantValue && (declaredType is null) ? Tokens.String : declaredType;
            _derivedType = DeriveTypeName(value, out bool derivedFromTypeHint);

            if ( _declaredType != null &&  _declaredType.Equals(Tokens.Date))
            {
                if (StringValueConverter.TryConvertString(AnnotateAsDateLiteral(ValueText), out _dateValue))
                {
                    ParsesToConstantValue = true;
                    ValueText = AnnotateAsDateLiteral(_dateValue.AsString);
                }
            }

            if (_derivedType.Equals(Tokens.Date))
            {
                if (StringValueConverter.TryConvertString(ValueText, out _dateValue))
                {
                    ValueText = AnnotateAsDateLiteral(_dateValue.AsString);
                    ParsesToConstantValue = true;
                }
            }

            if (derivedFromTypeHint)
            {
                _declaredType = _derivedType;
                ValueText = RemoveTypeHintChar(value);
            }
            else
            {
                ValueText = value.Replace("\"", "");
            }

            var conformToTypeName = _declaredType ?? _derivedType;
            ConformValueTextToType(conformToTypeName);
            TypeName = conformToTypeName;
        }

        private static bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");

        public decimal AsCurrency
        {
            get
            {
                if (this.TryConvertValue(out decimal result))
                {
                    return Math.Round(result, 4, MidpointRounding.ToEven);
                }
                throw new OverflowException();
            }
        }

        public string TypeName { get; set; }

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
                return _valueText;
            }
        }

        public bool ParsesToConstantValue { set; get; }

        public bool IsOverflowException { get; set; }

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

            var endingCharacter = inputValue.Last().ToString();
            return SymbolList.TypeHintToTypeName.ContainsKey(endingCharacter)
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

        private static string DeriveTypeName(string inputString, out bool derivedFromTypeHint)
        {
            derivedFromTypeHint = false;

            if (inputString.Length == 0)
            {
                return string.Empty;
            }

            if (TryParseAsDateLiteral(inputString, out ComparableDateValue _))
            {
                return Tokens.Date;
            }

            if (IsStringConstant(inputString))
            {
                return Tokens.String;
            }


            if (SymbolList.TypeHintToTypeName.TryGetValue(inputString.Last().ToString(), out string hintResult))
            {
                derivedFromTypeHint = true;
                return hintResult;
            }

            if (inputString.Contains(".") || inputString.Count(ch => ch.Equals('E')) == 1)
            {
                if (double.TryParse(inputString, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                {
                    return Tokens.Double;
                }
            }

            if (inputString.Equals(Tokens.True) || inputString.Equals(Tokens.False))
            {
                return Tokens.Boolean;
            }

            if (short.TryParse(inputString, out _))
            {
                return Tokens.Integer;
            }

            if (int.TryParse(inputString, out _))
            {
                return Tokens.Long;
            }

            if (TryParseAsHexLiteral(inputString, out short _))
            {
                return Tokens.Integer;
            }

            if (TryParseAsHexLiteral(inputString, out int _))
            {
                return Tokens.Long;
            }

            if (TryParseAsOctalLiteral(inputString, out short _))
            {
                return Tokens.Integer;
            }

            if (TryParseAsOctalLiteral(inputString, out int _))
            {
                return Tokens.Long;
            }

            if (long.TryParse(inputString, out _))
            {
                return Tokens.Double; //See 3.3.2 of the VBA specification.
            }

            return string.Empty;
        }

        private static bool TryParseAsDateLiteral(string valueString, out ComparableDateValue value)
        {
            return StringValueConverter.TryConvertString(valueString, out value);
        }

        private static bool TryParseAsHexLiteral(string valueString, out short value)
        {
            value = default;

            if (!valueString.StartsWith("&H") && !valueString.StartsWith("&h"))
            {
                return false;
            }

            var hexString = valueString.Substring(2).ToUpperInvariant();
            try
            {
                value = Convert.ToInt16(hexString, fromBase: 16);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        private static bool TryParseAsHexLiteral(string valueString, out int value)
        {
            value = default;

            if (!valueString.StartsWith("&H") && !valueString.StartsWith("&h"))
            {
                return false;
            }

            var hexString = valueString.Substring(2).ToUpperInvariant();
            try
            {
                value = Convert.ToInt32(hexString, fromBase: 16);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        private static bool TryParseAsHexLiteral(string valueString, out long value)
        {
            value = default;

            if (!valueString.StartsWith("&H") && !valueString.StartsWith("&h"))
            {
                return false;
            }

            var hexString = valueString.Substring(2).ToUpperInvariant();
            try
            {
                value = Convert.ToInt64(hexString, fromBase: 16);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        private static bool TryParseAsOctalLiteral(string valueString, out short value)
        {
            value = default;

            if (!valueString.StartsWith("&o") && !valueString.StartsWith("&O"))
            {
                return false;
            }

            var octalString = valueString.Substring(2);
            try
            {
                value = Convert.ToInt16(octalString, fromBase: 8);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        private static bool TryParseAsOctalLiteral(string valueString, out int value)
        {
            value = default;

            if (!valueString.StartsWith("&o") && !valueString.StartsWith("&O"))
            {
                return false;
            }

            var octalString = valueString.Substring(2);
            try
            {
                value = Convert.ToInt32(octalString, fromBase: 8);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        private static bool TryParseAsOctalLiteral(string valueString, out long value)
        {
            value = default;

            if (!valueString.StartsWith("&o") && !valueString.StartsWith("&O"))
            {
                return false;
            }

            var octalString = valueString.Substring(2);
            try
            {
                value = Convert.ToInt64(octalString, fromBase: 8);
                return true;
            }
            catch (OverflowException)
            {
                return false;
            }
        }

        private void ConformValueTextToType(string conformTypeName)
        {
            if (ValueText.Equals(double.NaN.ToString(CultureInfo.InvariantCulture)) &&
                !conformTypeName.Equals(Tokens.String))
            {
                return;
            }

            if (conformTypeName.Equals(Tokens.LongLong) || conformTypeName.Equals(Tokens.Long) ||
                conformTypeName.Equals(Tokens.Integer) || conformTypeName.Equals(Tokens.Byte))
            {
                if (this.TryConvertValue(out long newVal))
                {
                    ValueText = newVal.ToString();
                    ParsesToConstantValue = true;
                    CheckForOverflow(conformTypeName);
                    return;
                }

                if (conformTypeName.Equals(Tokens.Integer))
                {
                    if (TryParseAsHexLiteral(ValueText, out short outputHex))
                    {
                        ValueText = outputHex.ToString();
                        ParsesToConstantValue = true;
                        return;
                    }

                    if (TryParseAsOctalLiteral(ValueText, out short outputOctal))
                    {
                        ValueText = outputOctal.ToString();
                        ParsesToConstantValue = true;
                        return;
                    }
                }

                if (conformTypeName.Equals(Tokens.Long))
                {
                    if (TryParseAsHexLiteral(ValueText, out int outputHex))
                    {
                        ValueText = outputHex.ToString();
                        ParsesToConstantValue = true;
                        return;
                    }

                    if (TryParseAsOctalLiteral(ValueText, out int outputOctal))
                    {
                        ValueText = outputOctal.ToString();
                        ParsesToConstantValue = true;
                        return;
                    }
                }

                if (conformTypeName.Equals(Tokens.LongLong))
                {
                    if (TryParseAsHexLiteral(ValueText, out long outputHex))
                    {
                        ValueText = outputHex.ToString();
                        ParsesToConstantValue = true;
                        return;
                    }

                    if (TryParseAsOctalLiteral(ValueText, out long outputOctal))
                    {
                        ValueText = outputOctal.ToString();
                        ParsesToConstantValue = true;
                        return;
                    }
                }
            }

            if (conformTypeName.Equals(Tokens.Double) || conformTypeName.Equals(Tokens.Single))
            {
                if (this.TryConvertValue(out double newVal))
                {
                    ValueText = newVal.ToString(CultureInfo.InvariantCulture);
                    ParsesToConstantValue = true;
                    CheckForOverflow(conformTypeName);
                    return;
                }
            }

            if (conformTypeName.Equals(Tokens.Boolean))
            {
                if (this.TryConvertValue(out bool newVal))
                {
                    ValueText = newVal.ToString();
                    ParsesToConstantValue = true;
                    return;
                }
            }

            if (conformTypeName.Equals(Tokens.String))
            {
                ParsesToConstantValue = IsStringConstant(_inputValue);
                return;
            }

            if (conformTypeName.Equals(Tokens.Currency))
            {
                if (this.TryConvertValue(out decimal newVal))
                {
                    var currencyValue = Math.Round(newVal, 4, MidpointRounding.ToEven);
                    ValueText = currencyValue.ToString(CultureInfo.InvariantCulture);
                    ParsesToConstantValue = true;
                    CheckForOverflow(conformTypeName);
                    return;
                }
            }

            if (conformTypeName.Equals(Tokens.Date))
            {
                if (this.TryConvertValue(out double newVal))
                {
                    if (!_derivedType.Equals(Tokens.Date))
                    {
                        var dv = new DateValue(DateTime.FromOADate(newVal));
                        _dateValue = new ComparableDateValue(dv);
                        ValueText = _dateValue.AsDate.ToString(CultureInfo.InvariantCulture);
                        ParsesToConstantValue = true;
                    }
                }
            }
        }

        private static Dictionary<string, Action<string>> OverflowChecks = new Dictionary<string, Action<string>>()
        {
            [Tokens.Byte] = (a) => { byte.Parse(a); },
            [Tokens.Integer] = (a) => { Int16.Parse(a); },
            [Tokens.Long] = (a) => { Int32.Parse(a); },
            [Tokens.LongLong] = (a) => { Int64.Parse(a); },
            [Tokens.Single] = (a) => { float.Parse(a); },
            [Tokens.Currency] = (a) => { var value = decimal.Parse(a); if (value < CURRENCYMIN || value > CURRENCYMAX) { throw new OverflowException(); } },
        };

        private void CheckForOverflow(string typeName)
        {
            if (OverflowChecks.ContainsKey(typeName))
            {
                try
                {
                    OverflowChecks[typeName](ValueText);
                }
                catch (OverflowException)
                {
                    IsOverflowException = true;
                }
            }
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
