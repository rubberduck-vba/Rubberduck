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
        public const long LONGMIN = -2147486648;
        public const long LONGMAX = 2147486647;
        public const long INTEGERMIN = -32768;
        public const long INTEGERMAX = 32767;
        public const long BYTEMIN = 0;
        public const long BYTEMAX = 255;
        public const decimal CURRENCYMIN = -922337203685477.5808M;
        public const decimal CURRENCYMAX = 922337203685477.5807M;
        public const double SINGLEMIN = -3402823E38;
        public const double SINGLEMAX = 3402823E38;
    }

    public class SelectCaseInspectionValue
    {
        private readonly string _targetTypeName;
        private readonly string _valueAsString;
        private readonly Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool> _operatorIsGT;
        private readonly Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool> _operatorIsLT;
        private readonly Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool> _operatorIsEQ;

        private long? _valueAsLong;
        private double? _valueAsDouble;
        private decimal? _valueAsDecimal;
        private bool? _valueAsBoolean;

        private long resultLong;
        private double resultDouble;
        private decimal resultDecimal;

        private static Dictionary<string, Tuple<string, string>> TypeBoundaries = new Dictionary<string, Tuple<string, string>>()
        {
            { Tokens.Integer, new Tuple<string,string>(CompareExtents.INTEGERMIN.ToString(), CompareExtents.INTEGERMAX.ToString())},
            { Tokens.Long, new Tuple<string,string>(CompareExtents.LONGMIN.ToString(), CompareExtents.LONGMAX.ToString())},
            { Tokens.Byte, new Tuple<string,string>(CompareExtents.BYTEMIN.ToString(), CompareExtents.BYTEMAX.ToString())},
            { Tokens.Currency, new Tuple<string,string>(CompareExtents.CURRENCYMIN.ToString(), CompareExtents.CURRENCYMAX.ToString())},
            { Tokens.Single, new Tuple<string,string>(CompareExtents.SINGLEMIN.ToString(), CompareExtents.SINGLEMAX.ToString())}
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>> OperatorsIsGT = new Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsCurrency().Value > compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : !thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) > 0; } }
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>> OperatorsIsLT = new Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsCurrency().Value < compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) < 0; } }
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>> OperatorsIsEQ = new Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsCurrency().Value == compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) == 0; } }
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, bool>> IsParseableTests = new Dictionary<string, Func<SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsCurrency().HasValue; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsBoolean().HasValue; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue){ return true; } }
        };

        public SelectCaseInspectionValue(string valueToken, string targetTypeName)
        {
            _valueAsString = valueToken.EndsWith("#") ? valueToken.Replace("#", ".00") : valueToken;
            _targetTypeName = targetTypeName;

            Debug.Assert(OperatorsIsGT.ContainsKey(targetTypeName));
            Debug.Assert(OperatorsIsLT.ContainsKey(targetTypeName));
            Debug.Assert(OperatorsIsEQ.ContainsKey(targetTypeName));

            _operatorIsGT = OperatorsIsGT[targetTypeName];
            _operatorIsLT = OperatorsIsLT[targetTypeName];
            _operatorIsEQ = OperatorsIsEQ[targetTypeName];
        }

        public static SelectCaseInspectionValue CreateLowerBound(string typename)
        {
            if (TypeBoundaries.ContainsKey(typename))
            {
                return new SelectCaseInspectionValue(TypeBoundaries[typename].Item1, typename);
            }
            return null;
        }

        public static SelectCaseInspectionValue CreateUpperBound(string typename)
        {
            if (TypeBoundaries.ContainsKey(typename))
            {
                return new SelectCaseInspectionValue(TypeBoundaries[typename].Item2, typename);
            }
            return null;
        }

        public bool IsIntegerNumber => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte }.Contains(TargetTypeName);

        public string TargetTypeName => _targetTypeName;

        public bool IsParseable
            => IsParseableTests.ContainsKey(TargetTypeName) ? IsParseableTests[TargetTypeName](this) : false;

        public bool IsWithin(SelectCaseInspectionValue start, SelectCaseInspectionValue end ) 
            => start > end ? this >= end && this <= start : this >= start && this <= end;


        public static bool operator >(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue._operatorIsGT(thisValue, compValue);
        }

        public static bool operator <(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue._operatorIsLT(thisValue, compValue);
        }

        public static bool operator ==(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
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

        public static bool operator !=(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
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

        public static bool operator >=(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue == compValue || thisValue > compValue;
        }

        public static bool operator <=(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue == compValue || thisValue < compValue;
        }

        public override bool Equals(Object obj)
        {
            if (ReferenceEquals(null, obj) || !(obj is SelectCaseInspectionValue))
            {
                return false;
            }
            var asValue = (SelectCaseInspectionValue)obj;
            return asValue.TargetTypeName == TargetTypeName ? asValue == this : false;
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
                if (long.TryParse(_valueAsString, out resultLong))
                {
                    _valueAsLong = resultLong;
                }
                else if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsLong = SafeConvertToLong(resultDecimal);
                }
                else if (double.TryParse(_valueAsString, out resultDouble))
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

        public decimal? AsCurrency()
        {
            if (!_valueAsDecimal.HasValue)
            {
                if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsDecimal = resultDecimal;
                }
                else if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsDecimal = SafeConvertToDecimal(resultDouble);
                }
                else if (long.TryParse(_valueAsString, out resultLong))
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
                if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsDouble = resultDouble;
                }
                else if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsDouble = Convert.ToDouble(resultDecimal);
                }
                else if (long.TryParse(_valueAsString, out resultLong))
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
                else if (long.TryParse(_valueAsString, out resultLong))
                {
                    _valueAsBoolean = resultLong != 0;
                }
                else if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsBoolean = Math.Abs(resultDouble) > 0.00000001;
                }
                else if (decimal.TryParse(_valueAsString, out resultDecimal))
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
    }
}
