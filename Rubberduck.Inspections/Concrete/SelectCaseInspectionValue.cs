using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    internal static class CompareExtents
    {
        public const long LONGMIN = Int32.MinValue; //- 2147486648;
        public const long LONGMAX = Int32.MaxValue; //2147486647
        public const long INTEGERMIN = Int16.MinValue; //- 32768;
        public const long INTEGERMAX = Int16.MaxValue; //32767
        public const long BYTEMIN = byte.MinValue;  //0
        public const long BYTEMAX = byte.MaxValue;    //255
        public const decimal CURRENCYMIN = -922337203685477.5808M;
        public const decimal CURRENCYMAX = 922337203685477.5807M;
        public const double SINGLEMIN = -3402823E38;
        public const double SINGLEMAX = 3402823E38;
    }

    public class VBAValue
    {
        private readonly string _ctorTypeName;
        private readonly string _valueAsString;
        private string _useageTypeName;
        private  Func<VBAValue, VBAValue, bool> _operatorIsGT;
        private  Func<VBAValue, VBAValue, bool> _operatorIsLT;
        private  Func<VBAValue, VBAValue, bool> _operatorIsEQ;
        private  Func<VBAValue, VBAValue, VBAValue> _opMult;
        private  Func<VBAValue, VBAValue, VBAValue> _opDiv;
        private  Func<VBAValue, VBAValue, VBAValue> _opMinus;
        private  Func<VBAValue, VBAValue, VBAValue> _opPlus;

        private long? _valueAsLong;
        private long? _intValueAsLong;
        private int? _byteValueAsInt;
        private double? _valueAsDouble;
        private decimal? _valueAsDecimal;
        private long? _boolValueAsLong;

        private static Dictionary<string, Tuple<string, string>> TypeBoundaries = new Dictionary<string, Tuple<string, string>>()
        {
            { Tokens.Integer, new Tuple<string,string>(CompareExtents.INTEGERMIN.ToString(), CompareExtents.INTEGERMAX.ToString())},
            { Tokens.Long, new Tuple<string,string>(CompareExtents.LONGMIN.ToString(), CompareExtents.LONGMAX.ToString())},
            { Tokens.Byte, new Tuple<string,string>(CompareExtents.BYTEMIN.ToString(), CompareExtents.BYTEMAX.ToString())},
            { Tokens.Currency, new Tuple<string,string>(CompareExtents.CURRENCYMIN.ToString(), CompareExtents.CURRENCYMAX.ToString())},
            { Tokens.Single, new Tuple<string,string>(CompareExtents.SINGLEMIN.ToString(), CompareExtents.SINGLEMAX.ToString())}
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> OperatorsIsGT = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsInt().Value > compValue.AsInt().Value; } },
            { Tokens.Long, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsByte().Value > compValue.AsByte().Value; } },
            { Tokens.Double, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsCurrency().Value > compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : !thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) > 0; } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> OperatorsIsLT = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsInt().Value < compValue.AsInt().Value; } },
            { Tokens.Long, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsByte().Value < compValue.AsByte().Value; } },
            { Tokens.Double, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsCurrency().Value < compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) < 0; } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> OperatorsIsEQ = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsInt().Value == compValue.AsInt().Value; } },
            { Tokens.Long, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsByte().Value == compValue.AsByte().Value; } },
            { Tokens.Double, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsCurrency().Value == compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) == 0; } }
        };

        private static Dictionary<string, Func<VBAValue, bool>> HasValueTests = new Dictionary<string, Func<VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Long, delegate(VBAValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Byte, delegate(VBAValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Double, delegate(VBAValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Single, delegate(VBAValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Currency, delegate(VBAValue thisValue){ return thisValue.AsCurrency().HasValue; } },
            { Tokens.Boolean, delegate(VBAValue thisValue){ return thisValue.AsBoolean().HasValue; } },
            { Tokens.String, delegate(VBAValue thisValue){ return true; } }
        };

        private static Dictionary<string, Func<VBAValue, bool>> MaxMinTests = new Dictionary<string, Func<VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue){ return HasValueTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.INTEGERMAX) || (thisValue.AsLong() < CompareExtents.INTEGERMIN)  : false; } },
            { Tokens.Long, delegate(VBAValue thisValue){ return HasValueTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.LONGMAX) || (thisValue.AsLong() < CompareExtents.LONGMIN)  : false; } },
            { Tokens.Byte, delegate(VBAValue thisValue){ return HasValueTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.BYTEMAX) || (thisValue.AsLong() < CompareExtents.BYTEMIN)  : false; } },
            { Tokens.Double, delegate(VBAValue thisValue){ return false; } },
            { Tokens.Single, delegate(VBAValue thisValue){ return HasValueTests[Tokens.Single](thisValue) ? (thisValue.AsDouble() > CompareExtents.SINGLEMAX) || (thisValue.AsDouble() < CompareExtents.SINGLEMIN)  : false; } },
            { Tokens.Currency, delegate(VBAValue thisValue){ return HasValueTests[Tokens.Currency](thisValue) ? (thisValue.AsCurrency() > CompareExtents.CURRENCYMAX) || (thisValue.AsCurrency() < CompareExtents.CURRENCYMIN)  : false; } },
            { Tokens.Boolean, delegate(VBAValue thisValue){ return false; } },
            { Tokens.String, delegate(VBAValue thisValue){ return false; } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, VBAValue>> OperatorsMult = new Dictionary<string, Func<VBAValue, VBAValue, VBAValue>>()
        {
            { Tokens.Integer, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Long, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Byte, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Double, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value * RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Single, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value * RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Currency, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsCurrency().Value * RHS.AsCurrency()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Boolean, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, VBAValue>> OperatorsDiv = new Dictionary<string, Func<VBAValue, VBAValue, VBAValue>>()
        {
            { Tokens.Integer, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsInt().Value / RHS.AsInt().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Long, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value / RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Byte, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsByte().Value / RHS.AsByte().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Double, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value / RHS.AsDouble().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Single, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value / RHS.AsDouble().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Currency, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsCurrency().Value / RHS.AsCurrency().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Boolean, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value / RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, VBAValue>> OperatorsMinus = new Dictionary<string, Func<VBAValue, VBAValue, VBAValue>>()
        {
            { Tokens.Integer, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsInt().Value - RHS.AsInt().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Long, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value - RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Byte, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsByte().Value - RHS.AsByte().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Double, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value - RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Single, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value - RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Currency, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsCurrency().Value - RHS.AsCurrency()).Value.ToString(), LHS.UseageTypeName); } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, VBAValue>> OperatorsPlus = new Dictionary<string, Func<VBAValue, VBAValue, VBAValue>>()
        {
            { Tokens.Integer, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsInt().Value + RHS.AsInt().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Long, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsLong().Value + RHS.AsLong().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Byte, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsByte().Value + RHS.AsByte().Value).ToString(), LHS.UseageTypeName); } },
            { Tokens.Double, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value + RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Single, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsDouble().Value + RHS.AsDouble()).Value.ToString(), LHS.UseageTypeName); } },
            { Tokens.Currency, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue((LHS.AsCurrency().Value + RHS.AsCurrency()).Value.ToString(), LHS.UseageTypeName); } }
        };

        public VBAValue(string valueToken, string ctorTypeName)
        {
            _valueAsString = valueToken;
            var endingCharacter = _valueAsString.Last().ToString();
            if (new string[] { "#", "!", "@" }.Contains(endingCharacter) && !ctorTypeName.Equals(Tokens.String))
            {
                _valueAsString = _valueAsString.Replace(endingCharacter, ".00");
            }
            _ctorTypeName = ctorTypeName;
            UseageTypeName = _ctorTypeName;
        }

        public VBAValue(long value, string ctorTypeName = "Long")
        {
            _valueAsString = value.ToString();
            _ctorTypeName = ctorTypeName;
            UseageTypeName = _ctorTypeName;
        }

        public static string DeriveTypeName(string textValue, string defaultType = "String")
        {
            //TODO use TypeHintToTypeName - and add tests for each kind
            if (SymbolList.TypeHintToTypeName.TryGetValue(textValue.Last().ToString(), out string typeName))
            {
                return typeName;
            }

            if (textValue.StartsWith("\"") && textValue.EndsWith("\""))
            {
                return Tokens.String;
            }
            else if (textValue.Contains("."))
            {
                if (double.TryParse(textValue, out _))
                {
                    return Tokens.Double;
                }

                if (decimal.TryParse(textValue, out _))
                {
                    return Tokens.Currency;
                }
                return defaultType;
            }
            else if (textValue.Equals(Tokens.True) || textValue.Equals(Tokens.False))
            {
                return Tokens.Boolean;
            }
            else if (long.TryParse(textValue, out _))
            {
                return Tokens.Long;
            }
            else
            {
                return defaultType;
            }
        }

        public string OriginTypeName => _ctorTypeName;

        public string UseageTypeName
        {
            set
            {
                if(value != _useageTypeName)
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

        public bool HasValue
            => HasValueTests.ContainsKey(UseageTypeName) ? HasValueTests[UseageTypeName](this) : false;

        public bool IsWithin(VBAValue start, VBAValue end ) 
            => start > end ? this >= end && this <= start : this >= start && this <= end;

        public bool ExceedsMaxMin()
              => MaxMinTests.ContainsKey(UseageTypeName) ? MaxMinTests[UseageTypeName](this) : false;

        public static bool operator >(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._operatorIsGT(thisValue, compValue);
        }

        public static bool operator <(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._operatorIsLT(thisValue, compValue);
        }

        public static bool operator ==(VBAValue thisValue, VBAValue compValue)
        {
            if(ReferenceEquals(null, thisValue))
            {
                return ReferenceEquals(null, compValue);
            }
            else
            {
                return ReferenceEquals(null, compValue) ? false : thisValue._operatorIsEQ(thisValue, compValue);
            }
        }

        public static bool operator !=(VBAValue thisValue, VBAValue compValue)
        {
            if (ReferenceEquals(null, thisValue))
            {
                return !ReferenceEquals(null, compValue);
            }
            else
            {
                return ReferenceEquals(null, compValue) ? true : !(thisValue == compValue);
            }
        }

        public static bool operator >=(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue == compValue || thisValue > compValue;
        }

        public static bool operator <=(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue == compValue || thisValue < compValue;
        }

        public override bool Equals(Object obj)
        {
            if (ReferenceEquals(null, obj) || !(obj is VBAValue))
            {
                return false;
            }
            var asValue = (VBAValue)obj;
            return asValue.AsString() == AsString() ? asValue == this : false;
        }

        public static VBAValue operator *(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._opMult != null ? thisValue._opMult(thisValue, compValue) : null;
        }

        public static VBAValue operator /(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._opDiv != null ? thisValue._opDiv(thisValue, compValue) : null;
        }

        public static VBAValue operator -(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._opMinus != null ? thisValue._opMinus(thisValue, compValue) : null;
        }

        public static VBAValue operator +(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._opPlus != null ? thisValue._opPlus(thisValue, compValue) : null;
        }

        public static VBAValue operator ^(VBAValue thisValue, VBAValue compValue)
        {
            return (thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue)
                ? new VBAValue((Math.Pow(thisValue.AsDouble().Value, compValue.AsDouble().Value)).ToString(), thisValue.UseageTypeName)
                : null;
        }

        public static VBAValue operator %(VBAValue thisValue, VBAValue compValue)
        {
            return (thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue)
                ? new VBAValue((thisValue.AsDouble().Value % compValue.AsDouble().Value).ToString(), thisValue.UseageTypeName)
                : null;
        }

        public override int GetHashCode()
        {
            return _valueAsString.GetHashCode();
        }

        public string AsString()
        {
            return _valueAsString;
        }

        public override string ToString()
        {
            return AsString();
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
            if (_byteValueAsInt.HasValue && ( _byteValueAsInt.Value <= CompareExtents.BYTEMAX && _byteValueAsInt.Value >= CompareExtents.BYTEMIN))
            {
                byteResult = (byte)_byteValueAsInt.Value;
            }
            return byteResult;
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
            return _boolValueAsLong != 0 ? true : false;
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

        private Func<VBAValue, VBAValue, T> GetDelegate<T>(Dictionary<string, Func<VBAValue, VBAValue, T>> Operators, string targetTypeName)
        {
            return Operators.ContainsKey(targetTypeName) ? Operators[targetTypeName] : null;
        }
    }
}
