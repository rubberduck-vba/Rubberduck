using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionValue
    {
        string VbeText { set; get; }
        string ValueText { get; }
        string TypeName { set; get; }
        bool IsDeclaredTypeName { get; }
        bool IsConstantValue { get; set; }
    }

    public class UnreachableCaseInspectionValueConformed : IUnreachableCaseInspectionValue
    {
        private string _conformedValueText;
        private readonly IUnreachableCaseInspectionValue _decorated;
        private string _conformedTypeName;
        private bool _isDeclaredTypeName;
        public UnreachableCaseInspectionValueConformed(IUnreachableCaseInspectionValue value, string conformToTypeName)
        {
            _decorated = value;
            _conformedTypeName = value.TypeName;
            _isDeclaredTypeName = value.IsDeclaredTypeName && conformToTypeName.Equals(value.TypeName);
            if (!(conformToTypeName is null || conformToTypeName.Equals(string.Empty)))
            {
                ConformToType(conformToTypeName);
            }
            else
            {
                ConformToType(_decorated.TypeName);
            }
        }

        public string ValueText => _conformedValueText ?? _decorated.ValueText;

        public string VbeText { get => _decorated.VbeText; set => _decorated.VbeText = value; }
        public string TypeName { get => _conformedTypeName; set => _conformedTypeName = value; }

        public bool IsDeclaredTypeName => _isDeclaredTypeName;
        public bool IsConstantValue { get => _decorated.IsConstantValue; set => _decorated.IsConstantValue = value; }


        //TODO: this could be utility for both the base and decorated?
        private void ConformToType(string typeName)
        {
            if (typeName.Equals(Tokens.Long))
            {
                if (!long.TryParse(_decorated.ValueText, out _))
                {
                    if (double.TryParse(_decorated.ValueText, out double temp))
                    {
                        _conformedValueText = Convert.ToInt64(temp).ToString();
                        _conformedTypeName = typeName;
                        _decorated.IsConstantValue = true;
                    }
                }
                else
                {
                    if (long.TryParse(_decorated.ValueText, out long temp))
                    {
                        _conformedValueText = temp.ToString();
                        _conformedTypeName = typeName;
                        _decorated.IsConstantValue = true;
                    }
                    else
                    {
                        _decorated.IsConstantValue = false;
                    }
                }
            }
            else if (typeName.Equals(Tokens.Double))
            {
                if (double.TryParse(_decorated.ValueText, out double temp))
                {
                    _conformedValueText = temp.ToString();
                    _decorated.IsConstantValue = true;
                }
                else
                {
                    _decorated.IsConstantValue = false;
                }
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                if(_decorated.ValueText.Equals(Tokens.True) || _decorated.ValueText.Equals(Tokens.False))
                {
                    _decorated.IsConstantValue = true;
                }
                else if (double.TryParse(_decorated.ValueText, out double temp))
                {
                    _conformedValueText = temp != 0 ? Tokens.True : Tokens.False;
                    _decorated.IsConstantValue = true;
                }
                else
                {
                    _decorated.IsConstantValue = false;
                }
            }
        }
    }

    public class UnreachableCaseInspectionValue : IUnreachableCaseInspectionValue
    {
        private string _vbeText;
        private string _valueText;
        private string _declaredType;
        private string _assignedType;
        private bool _isConstantValue;

        public UnreachableCaseInspectionValue(string value)
        {
            _vbeText = value;
            if (IsStringConstant(value))
            {
                _isConstantValue = true;
                _declaredType = Tokens.String;
                _valueText = _vbeText.Replace("\"", "");
            }
            else
            {
                _isConstantValue = false;
                _declaredType = string.Empty;
                if (TryDeriveTypeName(_vbeText, out string typeName))
                {
                    _assignedType = typeName;
                }
                _valueText = RemoveTypeHintChar(_vbeText, _declaredType);
            }
        }

        public UnreachableCaseInspectionValue(double value)
        {
            _isConstantValue = true;
            _vbeText = value.ToString();
            _declaredType = Tokens.Double;
            _assignedType = string.Empty;

            _valueText = _vbeText;
        }

        public UnreachableCaseInspectionValue(decimal value)
        {
            _isConstantValue = true;
            _vbeText = value.ToString();
            _declaredType = Tokens.Currency;
            _assignedType = string.Empty;

            _valueText = _vbeText;
        }

        public UnreachableCaseInspectionValue(bool value)
        {
            _vbeText = value ? Tokens.True : Tokens.False;
            _isConstantValue = true;
            _declaredType = Tokens.Boolean;
            _assignedType = string.Empty;

            _valueText = _vbeText;
        }

        public UnreachableCaseInspectionValue(long value)
        {
            _vbeText = value.ToString();
            _isConstantValue = true;
            _declaredType = Tokens.Long;
            _assignedType = string.Empty;

            _valueText = _vbeText;
        }

        public UnreachableCaseInspectionValue(string value, string declaredType)
        {
            _vbeText = value;
            _assignedType = string.Empty;
            if ((declaredType is null) || declaredType.Equals(string.Empty))
            {
                _declaredType = string.Empty;
                if (TryDeriveTypeName(_vbeText, out string typeName))
                {
                    _assignedType = typeName;
                }
                _valueText = RemoveTypeHintChar(_vbeText, _declaredType);
            }
            else
            {
                _declaredType = declaredType;
            }

            _isConstantValue = IsStringConstant(value);

            _valueText = value.Replace("\"", "");
            if (!_declaredType.Equals(Tokens.String))
            {
                _valueText = RemoveTypeHintChar(_valueText, _declaredType);
            }
        }

        public string VbeText
        {
            set
            {
                if (_vbeText is null || _vbeText.Equals(string.Empty))
                {
                    _vbeText = value;
               }
            }
            get => _vbeText ?? string.Empty;
        }

        public string TypeName
        {
            set
            {
                if (_declaredType is null || _declaredType.Equals(string.Empty))
                {
                    _assignedType = value;
                }
            }

            get
            {
                if (_declaredType is null || _declaredType.Equals(string.Empty))
                {
                    if(_assignedType is null || _assignedType.Equals(string.Empty))
                    {
                        if (TryDeriveTypeName(VbeText, out string typeName))
                        {
                            TypeName = typeName;
                            return typeName;
                        }
                        return Tokens.Variant;
                    }
                    return _assignedType;
                }
                return _declaredType;
            }
        }

        public virtual string ValueText => _valueText;
        public bool IsDeclaredTypeName => !(_declaredType is null) && !(_declaredType.Equals(string.Empty));
        public bool IsConstantValue
        {
            get
            {
                return _isConstantValue;
            }
            set
            {
                _isConstantValue = value;
            }
        }

        private static string RemoveTypeHintChar(string inputValue, string declaredTypeName)
        {
            if (inputValue != string.Empty)
            {
                var endingCharacter = inputValue.Last().ToString();
                if (new string[] { "#", "!", "@" }.Contains(endingCharacter) && !declaredTypeName.Equals(Tokens.String))
                {
                    var regex = new Regex(@"^-*[0-9,\.]+$");
                    if (regex.IsMatch(inputValue.Replace(endingCharacter, "")))
                    {
                        if (inputValue.Contains("."))
                        {
                            //TODO: ???? - looks wrong
                            inputValue = inputValue.Replace(endingCharacter, "");
                        }
                        else
                        {
                            inputValue = inputValue.Replace(endingCharacter, ".00");
                        }
                    }
                }
            }
            return inputValue;
        }

        private static bool TryDeriveTypeName(string inputString, out string result)
        {
            result = string.Empty;
            if (inputString.Length == 0)
            {
                return false;
            }

            if (SymbolList.TypeHintToTypeName.TryGetValue(inputString.Last().ToString(), out result))
            {
                return true;
            }

            if (IsStringConstant(inputString))
            {
                result = Tokens.String;
                return true;
            }
            else if (inputString.Contains("."))
            {
                if (double.TryParse(inputString, out _))
                {
                    result = Tokens.Double;
                    return true;
                }

                if (decimal.TryParse(inputString, out _))
                {
                    result = Tokens.Currency;
                    return true;
                }
                return false;
            }
            else if (inputString.Count(ch => ch.Equals('E')) == 1)
            {
                if (double.TryParse(inputString, out _))
                {
                    result = Tokens.Double;
                    return true;
                }
                return false;
            }
            else if (inputString.Equals(Tokens.True) || inputString.Equals(Tokens.False))
            {
                result = Tokens.Boolean;
                return true;
            }
            else if (long.TryParse(inputString, out _))
            {
                result = Tokens.Long;
                return true;
            }
            return false;
        }

        public static List<string> IntegerTypes = new List<string>()
        {
            Tokens.Long,
            Tokens.Integer,
            Tokens.Byte
        };

        public static List<string> RationalTypes = new List<string>()
        {
            Tokens.Double,
            Tokens.Single,
            Tokens.Currency
        };

        private static bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");
    }
}
