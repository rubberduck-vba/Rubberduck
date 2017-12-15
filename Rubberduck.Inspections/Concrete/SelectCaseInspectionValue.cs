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

    public class VBEValue
    {
        private readonly string _targetTypeName;
        private readonly string _valueAsString;
        private readonly Func<VBEValue, VBEValue, bool> _operatorIsGT;
        private readonly Func<VBEValue, VBEValue, bool> _operatorIsLT;
        private readonly Func<VBEValue, VBEValue, bool> _operatorIsEQ;
        private readonly Func<VBEValue, VBEValue, VBEValue> _opMult;
        private readonly Func<VBEValue, VBEValue, VBEValue> _opDiv;
        private readonly Func<VBEValue, VBEValue, VBEValue> _opMinus;
        private readonly Func<VBEValue, VBEValue, VBEValue> _opPlus;

        private long? _valueAsLong;
        private long? _intValueAsLong;
        private int? _byteValueAsInt;
        private double? _valueAsDouble;
        private decimal? _valueAsDecimal;
        private bool? _valueAsBoolean;

        private static Dictionary<string, Tuple<string, string>> TypeBoundaries = new Dictionary<string, Tuple<string, string>>()
        {
            { Tokens.Integer, new Tuple<string,string>(CompareExtents.INTEGERMIN.ToString(), CompareExtents.INTEGERMAX.ToString())},
            { Tokens.Long, new Tuple<string,string>(CompareExtents.LONGMIN.ToString(), CompareExtents.LONGMAX.ToString())},
            { Tokens.Byte, new Tuple<string,string>(CompareExtents.BYTEMIN.ToString(), CompareExtents.BYTEMAX.ToString())},
            { Tokens.Currency, new Tuple<string,string>(CompareExtents.CURRENCYMIN.ToString(), CompareExtents.CURRENCYMAX.ToString())},
            { Tokens.Single, new Tuple<string,string>(CompareExtents.SINGLEMIN.ToString(), CompareExtents.SINGLEMAX.ToString())}
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, bool>> OperatorsIsGT = new Dictionary<string, Func<VBEValue, VBEValue, bool>>()
        {
            { Tokens.Integer, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsInt().Value > compValue.AsInt().Value; } },
            { Tokens.Long, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsByte().Value > compValue.AsByte().Value; } },
            { Tokens.Double, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsCurrency().Value > compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : !thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) > 0; } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, bool>> OperatorsIsLT = new Dictionary<string, Func<VBEValue, VBEValue, bool>>()
        {
            { Tokens.Integer, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsInt().Value < compValue.AsInt().Value; } },
            { Tokens.Long, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsByte().Value < compValue.AsByte().Value; } },
            { Tokens.Double, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsCurrency().Value < compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) < 0; } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, bool>> OperatorsIsEQ = new Dictionary<string, Func<VBEValue, VBEValue, bool>>()
        {
            { Tokens.Integer, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsInt().Value == compValue.AsInt().Value; } },
            { Tokens.Long, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsByte().Value == compValue.AsByte().Value; } },
            { Tokens.Double, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsCurrency().Value == compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBEValue thisValue, VBEValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) == 0; } }
        };

        private static Dictionary<string, Func<VBEValue, bool>> IsParseableTests = new Dictionary<string, Func<VBEValue, bool>>()
        {
            { Tokens.Integer, delegate(VBEValue thisValue){ return thisValue.AsInt().HasValue; } },
            { Tokens.Long, delegate(VBEValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Byte, delegate(VBEValue thisValue){ return thisValue.AsByte().HasValue; } },
            { Tokens.Double, delegate(VBEValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Single, delegate(VBEValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Currency, delegate(VBEValue thisValue){ return thisValue.AsCurrency().HasValue; } },
            { Tokens.Boolean, delegate(VBEValue thisValue){ return thisValue.AsBoolean().HasValue; } },
            { Tokens.String, delegate(VBEValue thisValue){ return true; } }
        };

        private static Dictionary<string, Func<VBEValue, bool>> MaxMinTests = new Dictionary<string, Func<VBEValue, bool>>()
        {
            { Tokens.Integer, delegate(VBEValue thisValue){ return IsParseableTests[Tokens.Integer](thisValue) ? (thisValue.AsLong() > CompareExtents.INTEGERMAX) || (thisValue.AsLong() < CompareExtents.INTEGERMIN)  : false; } },
            { Tokens.Long, delegate(VBEValue thisValue){ return IsParseableTests[Tokens.Long](thisValue) ? (thisValue.AsLong() > CompareExtents.LONGMAX) || (thisValue.AsLong() < CompareExtents.LONGMIN)  : false; } },
            { Tokens.Byte, delegate(VBEValue thisValue){ return IsParseableTests[Tokens.Byte](thisValue) ? (thisValue.AsLong() > CompareExtents.BYTEMAX) || (thisValue.AsLong() < CompareExtents.BYTEMIN)  : false; } },
            { Tokens.Double, delegate(VBEValue thisValue){ return false; } },
            { Tokens.Single, delegate(VBEValue thisValue){ return IsParseableTests[Tokens.Single](thisValue) ? (thisValue.AsDouble() > CompareExtents.SINGLEMAX) || (thisValue.AsDouble() < CompareExtents.SINGLEMIN)  : false; } },
            { Tokens.Currency, delegate(VBEValue thisValue){ return IsParseableTests[Tokens.Currency](thisValue) ? (thisValue.AsCurrency() > CompareExtents.CURRENCYMAX) || (thisValue.AsCurrency() < CompareExtents.CURRENCYMIN)  : false; } },
            { Tokens.Boolean, delegate(VBEValue thisValue){ return false; } },
            { Tokens.String, delegate(VBEValue thisValue){ return false; } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, VBEValue>> OperatorsMult = new Dictionary<string, Func<VBEValue, VBEValue, VBEValue>>()
        {
            { Tokens.Integer, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsInt().Value * RHS.AsInt().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Long, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsLong().Value * RHS.AsLong().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Byte, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsByte().Value * RHS.AsByte().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Double, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value * RHS.AsDouble()).Value.ToString(), LHS.TargetTypeName); } },
            { Tokens.Single, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value * RHS.AsDouble()).Value.ToString(), LHS.TargetTypeName); } },
            { Tokens.Currency, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsCurrency().Value * RHS.AsCurrency()).Value.ToString(), LHS.TargetTypeName); } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, VBEValue>> OperatorsDiv = new Dictionary<string, Func<VBEValue, VBEValue, VBEValue>>()
        {
            { Tokens.Integer, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsInt().Value / RHS.AsInt().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Long, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsLong().Value / RHS.AsLong().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Byte, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsByte().Value / RHS.AsByte().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Double, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value / RHS.AsDouble().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Single, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value / RHS.AsDouble().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Currency, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsCurrency().Value / RHS.AsCurrency().Value).ToString(), LHS.TargetTypeName); } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, VBEValue>> OperatorsMinus = new Dictionary<string, Func<VBEValue, VBEValue, VBEValue>>()
        {
            { Tokens.Integer, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsInt().Value - RHS.AsInt().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Long, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsLong().Value - RHS.AsLong().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Byte, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsByte().Value - RHS.AsByte().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Double, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value - RHS.AsDouble()).Value.ToString(), LHS.TargetTypeName); } },
            { Tokens.Single, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value - RHS.AsDouble()).Value.ToString(), LHS.TargetTypeName); } },
            { Tokens.Currency, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsCurrency().Value - RHS.AsCurrency()).Value.ToString(), LHS.TargetTypeName); } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, VBEValue>> OperatorsPlus = new Dictionary<string, Func<VBEValue, VBEValue, VBEValue>>()
        {
            { Tokens.Integer, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsInt().Value + RHS.AsInt().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Long, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsLong().Value + RHS.AsLong().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Byte, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsByte().Value + RHS.AsByte().Value).ToString(), LHS.TargetTypeName); } },
            { Tokens.Double, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value + RHS.AsDouble()).Value.ToString(), LHS.TargetTypeName); } },
            { Tokens.Single, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsDouble().Value + RHS.AsDouble()).Value.ToString(), LHS.TargetTypeName); } },
            { Tokens.Currency, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue((LHS.AsCurrency().Value + RHS.AsCurrency()).Value.ToString(), LHS.TargetTypeName); } }
        };

        public VBEValue(string valueToken, string targetTypeName)
        {
            _valueAsString = valueToken.EndsWith("#") && !targetTypeName.Equals(Tokens.String) ? valueToken.Replace("#", ".00") : valueToken;
            _targetTypeName = targetTypeName;

            _operatorIsGT = GetDelegate(OperatorsIsGT, targetTypeName);
            _operatorIsLT = GetDelegate(OperatorsIsLT, targetTypeName);
            _operatorIsEQ = GetDelegate(OperatorsIsEQ, targetTypeName);

            _opMult = GetDelegate(OperatorsMult, targetTypeName);
            _opDiv = GetDelegate(OperatorsDiv, targetTypeName);
            _opPlus = GetDelegate(OperatorsPlus, targetTypeName);
            _opMinus = GetDelegate(OperatorsMinus, targetTypeName);
        }

        public bool IsIntegerNumber => new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte }.Contains(TargetTypeName);

        public string TargetTypeName => _targetTypeName;

        public bool IsParseable
            => IsParseableTests.ContainsKey(TargetTypeName) ? IsParseableTests[TargetTypeName](this) : false;

        public bool IsWithin(VBEValue start, VBEValue end ) 
            => start > end ? this >= end && this <= start : this >= start && this <= end;

        public bool ExceedsMaxMin()
              => MaxMinTests.ContainsKey(TargetTypeName) ? MaxMinTests[TargetTypeName](this) : false;



        public static bool operator >(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue._operatorIsGT(thisValue, compValue);
        }

        public static bool operator <(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue._operatorIsLT(thisValue, compValue);
        }

        public static bool operator ==(VBEValue thisValue, VBEValue compValue)
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

        public static bool operator !=(VBEValue thisValue, VBEValue compValue)
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

        public static bool operator >=(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue == compValue || thisValue > compValue;
        }

        public static bool operator <=(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue == compValue || thisValue < compValue;
        }

        public override bool Equals(Object obj)
        {
            if (ReferenceEquals(null, obj) || !(obj is VBEValue))
            {
                return false;
            }
            var asValue = (VBEValue)obj;
            return asValue.TargetTypeName == TargetTypeName ? asValue == this : false;
        }

        public static VBEValue operator *(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue._opMult != null ? thisValue._opMult(thisValue, compValue) : null;
        }

        public static VBEValue operator /(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue._opDiv != null ? thisValue._opDiv(thisValue, compValue) : null;
        }

        public static VBEValue operator -(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue._opMinus != null ? thisValue._opMinus(thisValue, compValue) : null;
        }

        public static VBEValue operator +(VBEValue thisValue, VBEValue compValue)
        {
            return thisValue._opPlus != null ? thisValue._opPlus(thisValue, compValue) : null;
        }

        public static VBEValue operator ^(VBEValue thisValue, VBEValue compValue)
        {
            return (thisValue.AsDouble().HasValue && compValue.AsDouble().HasValue)
                ? new VBEValue((Math.Pow(thisValue.AsDouble().Value, compValue.AsDouble().Value)).ToString(), thisValue.TargetTypeName)
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
            if (_byteValueAsInt.HasValue && ( _byteValueAsInt.Value <= CompareExtents.BYTEMAX || _byteValueAsInt.Value >= CompareExtents.BYTEMIN))
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
            if (!_valueAsBoolean.HasValue)
            {
                if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsBoolean = true;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsBoolean = false;
                }
                else if (long.TryParse(_valueAsString, out long resultLong))
                {
                    _valueAsBoolean = resultLong != 0;
                }
                else if (double.TryParse(_valueAsString, out double resultDouble))
                {
                    _valueAsBoolean = Math.Abs(resultDouble) > 0.00000001;
                }
                else if (decimal.TryParse(_valueAsString, out decimal resultDecimal))
                {
                    _valueAsBoolean = Math.Abs(resultDecimal) > 0.0000001M;
                }
            }
            return _valueAsBoolean;
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

        private Func<VBEValue, VBEValue, T> GetDelegate<T>(Dictionary<string, Func<VBEValue, VBEValue, T>> Operators, string targetTypeName)
        {
            return Operators.ContainsKey(targetTypeName) ? Operators[targetTypeName] : null;
        }
    }
}
