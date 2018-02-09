using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Inspections.Concrete
{
    public static class CompareExtents
    {
        public static long LONGMIN = Int32.MinValue; //- 2147486648;
        public static long LONGMAX = Int32.MaxValue; //2147486647
        public static long INTEGERMIN = Int16.MinValue; //- 32768;
        public static long INTEGERMAX = Int16.MaxValue; //32767
        public static long BYTEMIN = byte.MinValue;  //0
        public static long BYTEMAX = byte.MaxValue;    //255
        public static decimal CURRENCYMIN = -922337203685477.5808M;
        public static decimal CURRENCYMAX = 922337203685477.5807M;
        public static double SINGLEMIN = -3402823E38;
        public static double SINGLEMAX = 3402823E38;
    }

    public class ParseTreeValue
    {
        private readonly string _valueAsString;
        private readonly string _inputString;
        private string _useageTypeName;
        private string _declaredTypeName;
        private string _derivedTypeName;
        private  Func<ParseTreeValue, ParseTreeValue, bool> _operatorIsGT;
        private  Func<ParseTreeValue, ParseTreeValue, bool> _operatorIsLT;
        private  Func<ParseTreeValue, ParseTreeValue, bool> _operatorIsEQ;
        private  Func<ParseTreeValue, ParseTreeValue, ParseTreeValue> _opMult;
        private  Func<ParseTreeValue, ParseTreeValue, ParseTreeValue> _opDiv;
        private  Func<ParseTreeValue, ParseTreeValue, ParseTreeValue> _opMinus;
        private  Func<ParseTreeValue, ParseTreeValue, ParseTreeValue> _opPlus;

        private long? _valueAsLong;
        private long? _intValueAsLong;
        private int? _byteValueAsInt;
        private double? _valueAsDouble;
        private decimal? _valueAsDecimal;
        private long? _boolValueAsLong;

        //private static Dictionary<string, Tuple<string, string>> TypeBoundaries = new Dictionary<string, Tuple<string, string>>()
        //{
        //    [Tokens.Integer] = new Tuple<string,string>(CompareExtents.INTEGERMIN.ToString(), CompareExtents.INTEGERMAX.ToString()),
        //    [Tokens.Long] = new Tuple<string,string>(CompareExtents.LONGMIN.ToString(), CompareExtents.LONGMAX.ToString()),
        //    [Tokens.Byte] = new Tuple<string,string>(CompareExtents.BYTEMIN.ToString(), CompareExtents.BYTEMAX.ToString()),
        //    [Tokens.Currency] = new Tuple<string,string>(CompareExtents.CURRENCYMIN.ToString(), CompareExtents.CURRENCYMAX.ToString()),
        //    [Tokens.Single] = new Tuple<string,string>(CompareExtents.SINGLEMIN.ToString(), CompareExtents.SINGLEMAX.ToString())
        //};

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, bool>> OperatorsIsGT = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, bool>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; },
            [Tokens.Long] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; },
            [Tokens.Byte] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; },
            [Tokens.Double] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; },
            [Tokens.Single] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; },
            [Tokens.Currency] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsCurrency().Value > compValue.AsCurrency().Value; },
            [Tokens.Boolean] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; },
            [Tokens.String] = delegate (ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) > 0; }
        };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, bool>> OperatorsIsLT = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, bool>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; },
            [Tokens.Long] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; },
            [Tokens.Byte] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; },
            [Tokens.Double] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; },
            [Tokens.Single] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; },
            [Tokens.Currency] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsCurrency().Value < compValue.AsCurrency().Value; },
            [Tokens.Boolean] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; },
            [Tokens.String] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) < 0; }
        };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, bool>> OperatorsIsEQ = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, bool>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().HasValue && compValue.AsInt().HasValue ? thisValue.AsLong().Value == compValue.AsInt().Value : false; },
            [Tokens.Long] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().HasValue && compValue.AsLong().HasValue ? thisValue.AsLong().Value == compValue.AsLong().Value : false; },
            [Tokens.Byte] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsLong().HasValue && compValue.AsByte().HasValue ? thisValue.AsLong().Value == compValue.AsByte().Value : false; },
            [Tokens.Double] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue ? thisValue.AsDouble().Value == compValue.AsDouble().Value : false; },
            [Tokens.Single] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue ? thisValue.AsDouble().Value == compValue.AsDouble().Value : false; },
            [Tokens.Currency] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsCurrency().HasValue && compValue.AsCurrency().HasValue ? thisValue.AsCurrency().Value == compValue.AsCurrency().Value : false; },
            [Tokens.Boolean] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsBoolean().HasValue && compValue.AsBoolean().HasValue ? thisValue.AsBoolean().Value == compValue.AsBoolean().Value : false; },
            [Tokens.String] = delegate(ParseTreeValue thisValue, ParseTreeValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) == 0; }
        };

        private static Dictionary<string, Func<ParseTreeValue, bool>> HasValueTests = new Dictionary<string, Func<ParseTreeValue, bool>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue thisValue){ return thisValue.AsLong().HasValue; },
            [Tokens.Long] = delegate(ParseTreeValue thisValue){ return thisValue.AsLong().HasValue; },
            [Tokens.Byte] = delegate(ParseTreeValue thisValue){ return thisValue.AsLong().HasValue; },
            [Tokens.Double] = delegate(ParseTreeValue thisValue){ return thisValue.AsDouble().HasValue; },
            [Tokens.Single] = delegate(ParseTreeValue thisValue){ return thisValue.AsDouble().HasValue; },
            [Tokens.Currency] = delegate(ParseTreeValue thisValue){ return thisValue.AsCurrency().HasValue; },
            [Tokens.Boolean] = delegate(ParseTreeValue thisValue){ return thisValue.AsBoolean().HasValue; },
            //[Tokens.String] = delegate (ParseTreeValue thisValue) { return true; }
            [Tokens.String] = delegate (ParseTreeValue thisValue) { return thisValue.InputStringIsStringConstant; }
        };

        private static Dictionary<string, Func<ParseTreeValue, bool>> MaxMinTests = new Dictionary<string, Func<ParseTreeValue, bool>>()
        {
            [Tokens.Integer] = delegate (ParseTreeValue thisValue) { return HasValueTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.INTEGERMAX) || (thisValue.AsLong() < CompareExtents.INTEGERMIN) : false; },
            [Tokens.Long] = delegate (ParseTreeValue thisValue) { return HasValueTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.LONGMAX) || (thisValue.AsLong() < CompareExtents.LONGMIN) : false; },
            [Tokens.Byte] = delegate (ParseTreeValue thisValue) { return HasValueTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.BYTEMAX) || (thisValue.AsLong() < CompareExtents.BYTEMIN) : false; },
            [Tokens.Double] = delegate (ParseTreeValue thisValue) { return false; },
            [Tokens.Single] = delegate (ParseTreeValue thisValue) { return HasValueTests[Tokens.Single](thisValue) ? (thisValue.AsDouble() > CompareExtents.SINGLEMAX) || (thisValue.AsDouble() < CompareExtents.SINGLEMIN) : false; },
            [Tokens.Currency] = delegate (ParseTreeValue thisValue) { return HasValueTests[Tokens.Currency](thisValue) ? (thisValue.AsCurrency() > CompareExtents.CURRENCYMAX) || (thisValue.AsCurrency() < CompareExtents.CURRENCYMIN) : false; },
            [Tokens.Boolean] = delegate (ParseTreeValue thisValue) { return false; },
            [Tokens.String] = delegate (ParseTreeValue thisValue) { return false; }
        };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>> OperatorsMult = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Long] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Byte] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Double] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value * RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); },
            [Tokens.Single] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value * RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); },
            [Tokens.Currency] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsCurrency().Value * RHS.AsCurrency()).Value.ToString(), LHS.UseageTypeName); },
            [Tokens.Boolean] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
        };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>> OperatorsDiv = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsInt().Value / RHS.AsInt().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Long] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value / RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Byte] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsByte().Value / RHS.AsByte().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Double] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value / RHS.AsDouble().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Single] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value / RHS.AsDouble().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Currency] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsCurrency().Value / RHS.AsCurrency().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Boolean] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value / RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
        };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>> OperatorsMinus = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
        {
            { Tokens.Integer, delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsInt().Value - RHS.AsInt().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Long, delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value - RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Byte, delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsByte().Value - RHS.AsByte().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Double, delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value - RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Single, delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value - RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Currency, delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsCurrency().Value - RHS.AsCurrency()).Value.ToString(), LHS.UseageTypeName); } }
        };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>> OperatorsPlus = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
        {
            [Tokens.Integer] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsInt().Value + RHS.AsInt().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Long] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsLong().Value + RHS.AsLong().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Byte] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsByte().Value + RHS.AsByte().Value).ToString(), LHS.UseageTypeName); },
            [Tokens.Double] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value + RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); },
            [Tokens.Single] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsDouble().Value + RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); },
            [Tokens.Currency] = delegate(ParseTreeValue LHS, ParseTreeValue RHS){ return new ParseTreeValue((LHS.AsCurrency().Value + RHS.AsCurrency()).Value.ToString(), LHS.UseageTypeName); }
        };

        public ParseTreeValue(string valueToken, string declaredTypeName = "")
        {
            _declaredTypeName = declaredTypeName ?? string.Empty;
            _useageTypeName = string.Empty;
            _inputString = valueToken;
            _valueAsString = _inputString.Replace("\"", "");
            if(_valueAsString != string.Empty)
            {
                var endingCharacter = _valueAsString.Last().ToString();
                if (new string[] { "#", "!", "@" }.Contains(endingCharacter) && !declaredTypeName.Equals(Tokens.String))
                {
                    var regex = new Regex(@"^-*[0-9,\.]+$");
                    if (regex.IsMatch(_valueAsString.Replace(endingCharacter, "")))
                    {
                        if (!_valueAsString.Contains("."))
                        {
                            _valueAsString = _valueAsString.Replace(endingCharacter, ".00");
                        }
                        else
                        {
                            _valueAsString = _valueAsString.Replace(endingCharacter, "");
                        }
                    }
                }
            }
            _derivedTypeName = string.Empty;

            if (HasValueTests.ContainsKey(declaredTypeName))
            {
                UseageTypeName = declaredTypeName;
            }
            else
            {
                UseageTypeName = DerivedTypeName;
            }
        }

        public ParseTreeValue(long value, string declaredTypeName = "")
        {
            _declaredTypeName = declaredTypeName ?? string.Empty;
            //_useageTypeName = string.Empty;
            _inputString = value.ToString();
            _valueAsString = _inputString;
            //UseageTypeName = declaredTypeName;
            _derivedTypeName = string.Empty;
            if (HasValueTests.ContainsKey(declaredTypeName))
            {
                UseageTypeName = declaredTypeName;
            }
            else
            {
                UseageTypeName = DerivedTypeName;
            }
            //if (!HasValueTests.ContainsKey(UseageTypeName))
            //{
            //    UseageTypeName = DerivedTypeName;
            //}
        }

        public static ParseTreeValue Null => new ParseTreeValue(string.Empty);
        public static ParseTreeValue Zero => new ParseTreeValue(0, Tokens.Long);
        public static ParseTreeValue Unity => new ParseTreeValue(1, Tokens.Long);
        public static ParseTreeValue False => Zero;
        public static ParseTreeValue True => new ParseTreeValue(-1, Tokens.Long);
        public ParseTreeValue AdditiveInverse => HasValue ? this * new ParseTreeValue(-1, UseageTypeName) : this;
        public static bool IsSupportedVBAType(string typeName) => OperatorsIsEQ.Keys.Contains(typeName);
        public static Byte MinValueByte => (Byte)CompareExtents.BYTEMIN;
        public static Byte MaxValueByte => (Byte)CompareExtents.BYTEMAX;

        private string DeriveTypeName(string defaultType = "String")
        {
            if(_inputString.Length == 0)
            {
                return Tokens.String;
            }

            if (SymbolList.TypeHintToTypeName.TryGetValue(_inputString.Last().ToString(), out string typeName))
            {
                return typeName;
            }

            if (InputStringIsStringConstant)
            {
                return Tokens.String;
            }
            else if (_inputString.Contains("."))
            {
                if (double.TryParse(_inputString, out _))
                {
                    return Tokens.Double;
                }

                if (decimal.TryParse(_inputString, out _))
                {
                    return Tokens.Currency;
                }
                return defaultType;
            }
            else if (_inputString.Equals(Tokens.True) || _inputString.Equals(Tokens.False))
            {
                return Tokens.Boolean;
            }
            else if (long.TryParse(_inputString, out _))
            {
                return Tokens.Long;
            }
            return defaultType;
        }

        //public static string DeriveTypeName(string inputValue, string defaultType = "String")
        //{
        //    if (SymbolList.TypeHintToTypeName.TryGetValue(inputValue.Last().ToString(), out string typeName))
        //    {
        //        return typeName;
        //    }

        //    if (IsStringConstant(inputValue))
        //    {
        //        return Tokens.String;
        //    }
        //    else if (inputValue.Contains("."))
        //    {
        //        if (double.TryParse(inputValue, out _))
        //        {
        //            return Tokens.Double;
        //        }

        //        if (decimal.TryParse(inputValue, out _))
        //        {
        //            return Tokens.Currency;
        //        }
        //        return defaultType;
        //    }
        //    else if (inputValue.Equals(Tokens.True) || inputValue.Equals(Tokens.False))
        //    {
        //        return Tokens.Boolean;
        //    }
        //    else if (long.TryParse(inputValue, out _))
        //    {
        //        return Tokens.Long;
        //    }
        //    return defaultType;
        //}
        public string DeclaredTypeName => _declaredTypeName;
        public bool HasDeclaredTypeName => !(_declaredTypeName is null) && !(_declaredTypeName.Equals(string.Empty));

        public string UseageTypeName
        {
            set
            {
                if (value != _useageTypeName)
                {
                     _useageTypeName = value;
                    _operatorIsGT = GetDelegate(OperatorsIsGT, _useageTypeName);
                    _operatorIsLT = GetDelegate(OperatorsIsLT, _useageTypeName);
                    _operatorIsEQ = GetDelegate(OperatorsIsEQ, _useageTypeName);
                    _opMult = GetDelegate(OperatorsMult, _useageTypeName);
                    _opDiv = GetDelegate(OperatorsDiv, _useageTypeName);
                    _opPlus = GetDelegate(OperatorsPlus, _useageTypeName);
                    _opMinus = GetDelegate(OperatorsMinus, _useageTypeName);
                }
            }
            get { return _useageTypeName; } 
        }

        public string DerivedTypeName
        {
            get
            {
                if (_derivedTypeName is null || _derivedTypeName.Equals(string.Empty))
                {
                    _derivedTypeName = DeriveTypeName(UseageTypeName);
                }
                return _derivedTypeName;
            }
        }

        public static bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");
        private bool InputStringIsStringConstant => IsStringConstant(_inputString);

        public bool HasValue
            => HasValueAs(UseageTypeName); // HasValueTests.ContainsKey(UseageTypeName) ? HasValueTests[UseageTypeName](this) : false;

        public bool HasValueAs(string typeName)
            => HasValueTests.ContainsKey(typeName) ? HasValueTests[typeName](this) : false;

        public bool IsWithin(ParseTreeValue start, ParseTreeValue end, bool isInclusive = true) 
            => isInclusive ?
                start > end ? this >= end && this <= start : this >= start && this <= end
                : start > end ? this > end && this < start : this > start && this < end;

        public bool ExceedsMaxMin()
              => MaxMinTests.ContainsKey(UseageTypeName) ? MaxMinTests[UseageTypeName](this) : false;


        public static bool operator >(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue._operatorIsGT(thisValue, compValue);
        }

        public static bool operator <(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue._operatorIsLT(thisValue, compValue);
        }

        public static bool operator ==(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            if (thisValue is null)
            {
                return (compValue is null);
            }
            if (!thisValue.UseageTypeName.Equals(Tokens.String) && thisValue.AsDouble().HasValue)
            {
                if (compValue.AsDouble().HasValue)
                {
                    return thisValue.AsDouble().Value.CompareTo(compValue.AsDouble().Value) == 0;
                }
                return false;
            }
            return thisValue.AsString().Equals(compValue.AsString());
        }

        public static bool operator !=(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            if (thisValue is null)
            {
                return !(compValue is null);
            }
            return compValue is null ? true : !(thisValue == compValue);
        }

        public static bool operator >=(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue == compValue || thisValue > compValue;
        }

        public static bool operator <=(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue == compValue || thisValue < compValue;
        }

        public override bool Equals(Object obj)
        {
            if (obj is null || !(obj is ParseTreeValue))
            {
                return false;
            }
            var asValue = (ParseTreeValue)obj;
            return asValue == this;
        }

        public static ParseTreeValue operator *(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue._opMult != null ? thisValue._opMult(thisValue, compValue) : null;
        }

        public static ParseTreeValue operator /(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue._opDiv != null ? thisValue._opDiv(thisValue, compValue) : null;
        }

        public static ParseTreeValue operator -(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue._opMinus != null ? thisValue._opMinus(thisValue, compValue) : null;
        }

        public static ParseTreeValue operator +(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return thisValue._opPlus != null ? thisValue._opPlus(thisValue, compValue) : null;
        }

        public static ParseTreeValue Pow(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return (thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue)
                ? new ParseTreeValue((Math.Pow(thisValue.AsDouble().Value, compValue.AsDouble().Value)).ToString(), thisValue.UseageTypeName)
                : null;
        }

        public static ParseTreeValue operator !(ParseTreeValue thisValue)
        {
            return (thisValue.AsBoolean().HasValue)
                ? new ParseTreeValue((!thisValue.AsBoolean().Value).ToString(), thisValue.UseageTypeName)
                : null;
        }

        public static ParseTreeValue operator %(ParseTreeValue thisValue, ParseTreeValue compValue)
        {
            return (thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue)
                ? new ParseTreeValue((thisValue.AsDouble().Value % compValue.AsDouble().Value).ToString(), thisValue.UseageTypeName)
                : null;
        }

        public override int GetHashCode()
        {
            if (!UseageTypeName.Equals(Tokens.String) && AsDouble().HasValue)
            {
                return AsDouble().Value.GetHashCode();
            }
            return _valueAsString.GetHashCode();
        }

        public string AsString()
        {
            return InputStringIsStringConstant ? _inputString : _valueAsString;
        }

        public override string ToString()
        {
            return AsString();
        }

        public static explicit operator long(ParseTreeValue thisVal)
        {
            if(thisVal.TryGetValue(out long v))
            {
                return v;
            }
            return default;
        }

        public bool TryGetValue(out long v)
        {
            v = 0;
            if (AsLong().HasValue)
            {
                v = AsLong().Value;
                return true;
            }
            return false;
        }

        public long? AsLong()
        {
            if (!_valueAsLong.HasValue)
            {
                if (long.TryParse(_valueAsString, out long resultLong))
                {
                    _valueAsLong = resultLong;
                }
                else if (decimal.TryParse(_valueAsString, out decimal resultDecimal))
                {
                    _valueAsLong = SafeConvertToLong(resultDecimal);
                }
                else if (double.TryParse(_valueAsString, out double resultDouble))
                {
                    _valueAsLong = SafeConvertToLong(resultDouble);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsLong = -1;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsLong = 0;
                }
            }
            return _valueAsLong;
        }

        public bool TryGetValue(out int v)
        {
            v = 0;
            if (AsInt().HasValue)
            {
                v = AsInt().Value;
                return true;
            }
            return false;
        }

        public int? AsInt()
        {
            if (!_intValueAsLong.HasValue)
            {
                if (long.TryParse(_valueAsString, out long result))
                {
                    _intValueAsLong = result;
                }
                else if (decimal.TryParse(_valueAsString, out decimal resultDec))
                {
                    _intValueAsLong = SafeConvertToInteger(resultDec);
                }
                else if (double.TryParse(_valueAsString, out double resultDbl))
                {
                    _intValueAsLong = SafeConvertToInteger(resultDbl);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _intValueAsLong = -1;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _intValueAsLong = 0;
                }
            }

            int? intResult = null;
            if (_intValueAsLong.HasValue && (_intValueAsLong.Value <= CompareExtents.INTEGERMAX || _intValueAsLong.Value >= CompareExtents.INTEGERMIN))
            {
                intResult = (int)_intValueAsLong.Value;
            }
            return intResult;
        }

        public bool TryGetValue(out byte v)
        {
            v = 0;
            if (AsByte().HasValue)
            {
                v = AsByte().Value;
                return true;
            }
            return false;
        }

        public byte? AsByte()
        {
            if (!_byteValueAsInt.HasValue)
            {
                if (int.TryParse(_valueAsString, out int result))
                {
                    _byteValueAsInt = result;
                }
                else if (decimal.TryParse(_valueAsString, out decimal resultDecimal))
                {
                    _byteValueAsInt = SafeConvertToByte(resultDecimal);
                }
                else if (double.TryParse(_valueAsString, out double resultDouble))
                {
                    _byteValueAsInt = SafeConvertToByte(resultDouble);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _byteValueAsInt = 1;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _byteValueAsInt = 0;
                }
            }

            byte? byteResult = null;
            if (_byteValueAsInt.HasValue && (_byteValueAsInt.Value <= CompareExtents.BYTEMAX && _byteValueAsInt.Value >= CompareExtents.BYTEMIN))
            {
                byteResult = (byte)_byteValueAsInt.Value;
            }
            return byteResult;
        }

        public bool TryGetValue(out decimal v)
        {
            v = 0.0M;
            if (AsCurrency().HasValue)
            {
                v = AsCurrency().Value;
                return true;
            }
            return false;
        }

        public decimal? AsCurrency()
        {
            if (!_valueAsDecimal.HasValue)
            {
                if (decimal.TryParse(_valueAsString, out decimal resultDecimal))
                {
                    _valueAsDecimal = resultDecimal;
                }
                else if (double.TryParse(_valueAsString, out double resultDouble))
                {
                    _valueAsDecimal = SafeConvertToDecimal(resultDouble);
                }
                else if (long.TryParse(_valueAsString, out long resultLong))
                {
                    _valueAsDecimal = SafeConvertToDecimal(resultLong);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsDecimal = -1.0M;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsDecimal = 0.0M;
                }
            }
            return _valueAsDecimal;
        }

        public bool TryGetValue(out double v)
        {
            v = 0.0;
            if (AsDouble().HasValue)
            {
                v = AsDouble().Value;
                return true;
            }
            return false;
        }

        public double? AsDouble()
        {
            if (!_valueAsDouble.HasValue)
            {
                if (double.TryParse(_valueAsString, out double resultDouble))
                {
                    _valueAsDouble = resultDouble;
                }
                else if (decimal.TryParse(_valueAsString, out decimal resultDecimal))
                {
                    _valueAsDouble = Convert.ToDouble(resultDecimal);
                }
                else if (long.TryParse(_valueAsString, out long resultLong))
                {
                    _valueAsDouble = Convert.ToDouble(resultLong);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsDouble = -1.0;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsDouble = 0.0;
                }
            }
            return _valueAsDouble;
        }

        public bool TryGetValue(out bool v)
        {
            v = false;
            if (AsBoolean().HasValue)
            {
                v = AsBoolean().Value;
                return true;
            }
            return false;
        }

        public bool? AsBoolean()
        {
            if (!_boolValueAsLong.HasValue)
            {
                if (_valueAsString.Equals(Tokens.True))
                {
                    _boolValueAsLong = -1;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _boolValueAsLong = 0;
                }
                else if (long.TryParse(_valueAsString, out long resultLong))
                {
                    _boolValueAsLong = resultLong;
                }
                else if (double.TryParse(_valueAsString, out double resultDouble))
                {
                    _boolValueAsLong = Math.Abs(resultDouble) > 0.00000001 ? -1 : 0;
                }
                else if (decimal.TryParse(_valueAsString, out decimal resultDecimal))
                {
                    _boolValueAsLong = Math.Abs(resultDecimal) > 0.0000001M ? -1 : 0;
                }
            }
            if (_boolValueAsLong == null)
            {
                return null;
            }
            return _boolValueAsLong != 0 ? true : false;
        }

        public bool TryGetValue(out string v)
        {
            v = string.Empty;
            if (!AsString().Equals(string.Empty))
            {
                v = AsString();
                return true;
            }
            return false;
        }

        private long? SafeConvertToLong<T>(T value)
        {
            try
            {
                return Convert.ToInt64(value);
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private int? SafeConvertToInteger<T>(T value)
        {
            try
            {
                return Convert.ToInt32(value);
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private byte? SafeConvertToByte<T>(T value)
        {
            try
            {
                return Convert.ToByte(value);
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private decimal? SafeConvertToDecimal<T>(T value)
        {
            try
            {
                return Convert.ToDecimal(value);
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private Func<ParseTreeValue, ParseTreeValue, T> GetDelegate<T>(Dictionary<string, Func<ParseTreeValue, ParseTreeValue, T>> Operators, string targetTypeName)
        {
            return Operators.ContainsKey(targetTypeName) ? Operators[targetTypeName] : null;
        }
    }
}
