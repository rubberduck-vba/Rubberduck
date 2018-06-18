using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValue
    {
        string ValueText { get; }
        string TypeName { get; }
        bool ParsesToConstantValue { get; set; }
    }

    public class ParseTreeValue : IParseTreeValue
    {
        private readonly string _inputValue;
        private readonly string _declaredType;
        private readonly string _derivedType;

        public ParseTreeValue(string value, string declaredType = null, string conformToType = null)
        {
            if (value is null)
            {
                throw new ArgumentNullException("null 'value' argument passed to UCIValue");
            }
            _inputValue = value;
            ValueText = value;
            ParsesToConstantValue = IsStringConstant(value);
            _declaredType = ParsesToConstantValue && (declaredType is null) ? Tokens.String : declaredType;
            _derivedType = DeriveTypeName(value, out bool derivedFromTypeHint);
            if (derivedFromTypeHint)
            {
                _declaredType = _derivedType;
                ValueText = RemoveTypeHintChar(value);
            }
            else
            {
                ValueText = value.Replace("\"", "");
            }

            var conformToTypeName = conformToType ?? _declaredType ?? _derivedType;
            ConformValueTextToType(conformToTypeName);
        }

        private static bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");

        public string TypeName => _declaredType ?? _derivedType ?? string.Empty;

        public string ValueText { private set; get; }

        public bool ParsesToConstantValue { set; get; }

        public override string ToString() => ValueText;

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

        private static string DeriveTypeName(string inputString, out bool derivedFromTypeHint)
        {
            derivedFromTypeHint = false;

            if (inputString.Length == 0)
            {
                return string.Empty;
            }

            if (SymbolList.TypeHintToTypeName.TryGetValue(inputString.Last().ToString(), out string hintResult))
            {
                derivedFromTypeHint = true;
                return hintResult;
            }

            if (IsStringConstant(inputString))
            {
                return Tokens.String;
            }

            if (inputString.Contains("."))
            {
                if (double.TryParse(inputString, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                {
                    return Tokens.Double;
                }
            }

            if (inputString.Count(ch => ch.Equals('E')) == 1)
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
                    ValueText = newVal.ToString(CultureInfo.InvariantCulture);
                    ParsesToConstantValue = true;
                    return;
                }
            }
        }
    }

    public static class ParseTreeValueExtensions
    {
        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out long value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value, parseTreeValue.TypeName);

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out double value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out decimal value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out bool value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);

        public static bool TryConvertValue(this IParseTreeValue parseTreeValue, out string value)
            => StringValueConverter.TryConvertString(parseTreeValue.ValueText, out value);
    }
}
