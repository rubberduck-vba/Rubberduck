using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    public class RangeClause : IComparable, IRangeClause
    {
        private const string EQ = "=";
        private const string NEQ = "<>";
        private const string LT = "<";
        private const string LTE = "<=";
        private const string GT = ">";
        private const string GTE = ">=";

        private readonly VBAParser.RangeClauseContext _ctxt;
        private readonly QualifiedContext<ParserRuleContext> _qualifiedCaseContext;
        private readonly string _typeName;
        private bool _usesIsClause;
        private KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext> _rangeContexts;
        private bool _isRange;
        private readonly bool _isSingleVal;
        private string _compareSymbol;
        public bool IsPartiallyEquivalent { get; set; }
        public bool IsFullyEquivalent { get; set; }
        private Dictionary<string, Func<string, string, string, int>> _singleValueCompares = new Dictionary<string, Func<string, string, string, int>>();
        private Dictionary<string, Func<string, string, string, string, int>> _isStmtToIsStmtCompares = new Dictionary<string, Func<string, string, string, string, int>>();
        private Dictionary<string, Func<string, string, string, int>> _isWithinCompares = new Dictionary<string, Func<string, string, string, int>>();
        private Dictionary<string, Func<string, string, string, int>> _singleValueToIsStmtComparers = new Dictionary<string, Func<string, string, string, int>>();
        private Dictionary<string, Func<string, string, string, string, int>> _rangeToIsStmtCompares = new Dictionary<string, Func<string, string, string, string, int>>();

        private static string[] LongComparisonTypes = { "Integer", "Long", "Byte" };
        private static string[] DoubleComparisonTypes = { "Double", "Single" };
        private static string[] CurrencyComparisonTypes = { "Currency" };
        private static string[] BooleanComparisonTypes = { "Boolean" };

        private static Dictionary<string, string> _comparisonOperatorsAndInversions = new Dictionary<string, string>()
        {
            { EQ,EQ },
            { NEQ,NEQ },
            { LT,GT },
            { LTE,GTE },
            { GT,LT },
            { GTE,LTE }
        };

        public RangeClause(QualifiedContext<ParserRuleContext> caseClause, VBAParser.RangeClauseContext ctxt, IdentifierReference theRef, string typeName)
        {
            _qualifiedCaseContext = caseClause;
            _ctxt = ctxt;
            _theRef = theRef;
            _typeName = typeName;
            _compareSymbol = DetermineTheComparisonOperator(ctxt);
            _usesIsClause = HasChildToken(ctxt, Tokens.Is);
            _isRange = HasChildToken(ctxt, Tokens.To);
            _isSingleVal = true;
            IsParseable = true;

            if (_isRange)
            {
                _rangeContexts = new KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext>
                    (ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(ctxt),
                    ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(ctxt));
                IsParseable = _rangeContexts.Key != null && _rangeContexts.Value != null;
                _isSingleVal = false;
            }
            SetIsParseable();
        }

        public string ClauseTypeName => _typeName;
        public bool IsSingleVal => _isSingleVal;
        public bool UsesIsClause => _usesIsClause;
        public bool IsRange => _isRange;
        public string CompareSymbol => _compareSymbol;
        public string TypeName => _typeName;
        public bool IsParseable { get; set; }
        private bool IsComparisonOperator(string opCandidate) { return _comparisonOperatorsAndInversions.Keys.Contains(opCandidate); }
        private string InvertComparisonOperator(string theOperator)
        {
            return IsComparisonOperator(theOperator) ? _comparisonOperatorsAndInversions[theOperator] : theOperator;
        }

        private IdentifierReference _theRef;

        private bool isLongType => LongComparisonTypes.Contains(_typeName);
        private bool isDoubleType => DoubleComparisonTypes.Contains(_typeName);
        private bool isBooleanType => BooleanComparisonTypes.Contains(_typeName);
        private bool isDecimalType => CurrencyComparisonTypes.Contains(_typeName);

        public string ValueAsString => GetRangeClauseText(_ctxt);
        public string ValueMinAsString => _isRange ? _rangeContexts.Key.GetText() : ValueAsString;
        public string ValueMaxAsString => _isRange ? _rangeContexts.Value.GetText() : ValueAsString;


        private string GetRangeClauseText(VBAParser.RangeClauseContext ctxt)
        {
            VBAParser.RelationalOpContext relationalOpCtxt;
            if (TryGetExprContext(ctxt, out relationalOpCtxt))
            {
                return GetTextForRelationalOpContext(relationalOpCtxt);
            }

            VBAParser.UnaryMinusOpContext negativeCtxt;
            if (TryGetExprContext(ctxt, out negativeCtxt))
            {
                return negativeCtxt.GetText();
            }
            else
            {
                VBAParser.LiteralExprContext theValCtxt;
                return TryGetExprContext(ctxt, out theValCtxt) ? theValCtxt.GetText() : string.Empty;
            }
        }

        private string GetTextForRelationalOpContext(VBAParser.RelationalOpContext relationalOpCtxt)
        {
            VBAParser.LExprContext lExprCtxt = null;
            VBAParser.LiteralExprContext theValueCtxt = null;

            var lExprCtxtIndex = -1;
            var literalExprCtxtIndex = -1;
            for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
            {
                var text = relationalOpCtxt.children[idx].GetText();
                if (relationalOpCtxt.children[idx] is VBAParser.LExprContext)
                {
                    lExprCtxt = (VBAParser.LExprContext)relationalOpCtxt.children[idx];
                    lExprCtxtIndex = idx;
                }
                else if (relationalOpCtxt.children[idx] is VBAParser.LiteralExprContext)
                {
                    theValueCtxt = (VBAParser.LiteralExprContext)relationalOpCtxt.children[idx];
                    literalExprCtxtIndex = idx;
                }
                else if (IsComparisonOperator(text))
                {
                    _compareSymbol = text;
                }
            }

            if (lExprCtxtIndex > literalExprCtxtIndex)
            {
                _compareSymbol = _comparisonOperatorsAndInversions[_compareSymbol];
            }

            if (lExprCtxt.GetText().Equals(_theRef.IdentifierName))
            {
                //If 'z' is the Select Case variable, 
                //then 'z < 10' will be treated as 'Is < 10'
                //and '10 < z' will be treated as 'Is > 10
                _usesIsClause = true;
                return theValueCtxt.GetText();
            }
            return string.Empty;
        }

        private bool StringToBool(string strValue)
        {
            if (strValue.Equals("True"))
            {
                return true;
            }
            else if (strValue.Equals("False"))
            {
                return false;
            }
            else
            {
                return int.Parse(strValue) != 0;
            }
        }

        private Func<string, string, string, int> GetSingleValueComparer()
        {
            if (!_singleValueCompares.Any())
            {
                _singleValueCompares.Add("Long", new Func<string, string, string, int>(CompareSingleValuesLong));
                _singleValueCompares.Add("Integer", new Func<string, string, string, int>(CompareSingleValuesLong));
                _singleValueCompares.Add("Byte", new Func<string, string, string, int>(CompareSingleValuesLong));
                _singleValueCompares.Add("Double", new Func<string, string, string, int>(CompareSingleValuesDouble));
                _singleValueCompares.Add("Single", new Func<string, string, string, int>(CompareSingleValuesDouble));
                _singleValueCompares.Add("Boolean", new Func<string, string, string, int>(CompareSingleValuesBoolean));
                _singleValueCompares.Add("Currency", new Func<string, string, string, int>(CompareSingleValuesDecimal));
            }

            Func<string, string, string, int> comparer;

            if (!_singleValueCompares.TryGetValue(_typeName, out comparer))
            {
                comparer = CompareSingleValues;
            }
            return comparer;
        }

        public int CompareTo(object obj)
        {
            var prior = obj as IRangeClause;

            if(!_typeName.Equals(prior.TypeName))
            {
                //Inconsistent type => different inspection?
                return 0;
            }

            if (IsSingleVal && prior.IsSingleVal)
            {
                var comparer = GetSingleValueComparer();

                if (!UsesIsClause && !prior.UsesIsClause)
                {
                    var result = comparer(prior.ValueAsString, ValueAsString, EQ);

                    IsFullyEquivalent = result == 0;
                    return result;
                }
                else if (!UsesIsClause && prior.UsesIsClause)
                {
                    var result = comparer(prior.ValueAsString, ValueAsString, prior.CompareSymbol);

                    IsFullyEquivalent = (result == 0);
                    return result;
                }
                else if (UsesIsClause && !prior.UsesIsClause)
                {
                    var result = comparer(ValueAsString, prior.ValueAsString, CompareSymbol);

                    IsPartiallyEquivalent = (result == 0);
                    return result;
                }
                else if (UsesIsClause && prior.UsesIsClause)
                {
                    var isStmtToIsStmtComparer = GetIsStmtComparer();
                    return isStmtToIsStmtComparer(prior.ValueAsString, ValueAsString, prior.CompareSymbol, CompareSymbol);
                }
            }
            else if (IsSingleVal && prior.IsRange)
            {
                if (!UsesIsClause)
                {
                    var isWithin = GetIsWithinComparer();

                    var result = isWithin(ValueAsString, prior.ValueMinAsString, prior.ValueMaxAsString);
                    IsFullyEquivalent = result == 0;
                    return result;
                }
                else
                {
                    // e.g. Case Is > 8 prior Case 3 to 10
                    var comparer = GetSingleValueToIsStmtComparer();

                    var resultStartVal = comparer(ValueAsString, prior.ValueMinAsString, CompareSymbol);
                    var resultEndVal = comparer(ValueAsString, prior.ValueMaxAsString, CompareSymbol);

                    return resultStartVal == 0 || resultEndVal == 0 ? 0 : 1;
                }
            }
            else if (IsRange && prior.IsSingleVal)
            {
                if (!prior.UsesIsClause)
                {
                    var comparer = GetIsWithinComparer();

                    var result = comparer(prior.ValueAsString, ValueMinAsString, ValueMaxAsString);
                    IsPartiallyEquivalent = result == 0;
                    return result;
                }
                else
                {
                    var compareRangeToIsStmt = GetRangeToIsStmtComparer();

                    return compareRangeToIsStmt(prior.ValueMinAsString, ValueMinAsString, ValueMaxAsString, prior.CompareSymbol);
                }
            }
            else if (IsRange && prior.IsRange)
            {
                var isWithin = GetIsWithinComparer();

                if (isWithin(ValueMinAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0
                        && isWithin(ValueMaxAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0)
                {
                    IsFullyEquivalent = true;
                    return 0;
                }
                else
                {
                    IsPartiallyEquivalent = isWithin(ValueMinAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0
                        || isWithin(ValueMaxAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0;

                    return IsPartiallyEquivalent ? 0 : ValueMaxAsString.CompareTo(prior.ValueMaxAsString);
                }
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
        }

        private Func<string, string, string, int> GetSingleValueToIsStmtComparer()
        {
            if (!_singleValueToIsStmtComparers.Any())
            {
                _singleValueToIsStmtComparers.Add("Long", new Func<string, string, string, int>(CompareSingleValueToIsStmtLong));
                _singleValueToIsStmtComparers.Add("Integer", new Func<string, string, string, int>(CompareSingleValueToIsStmtLong));
                _singleValueToIsStmtComparers.Add("Byte", new Func<string, string, string, int>(CompareSingleValueToIsStmtLong));
                _singleValueToIsStmtComparers.Add("Double", new Func<string, string, string, int>(CompareSingleValueToIsStmtDouble));
                _singleValueToIsStmtComparers.Add("Single", new Func<string, string, string, int>(CompareSingleValueToIsStmtDouble));
                _singleValueToIsStmtComparers.Add("Boolean", new Func<string, string, string, int>(CompareSingleValueToIsStmtBoolean));
                _singleValueToIsStmtComparers.Add("Currency", new Func<string, string, string, int>(CompareSingleValueToIsStmtDecimal));
            }

            Func<string, string, string, int> comparer;

            if (!_singleValueToIsStmtComparers.TryGetValue(_typeName, out comparer))
            {
                comparer = CompareSingleValueToIsStmt;
            }
            return comparer;
        }

        private Func<string, string, string, int> GetIsWithinComparer()
        {
            if (!_isWithinCompares.Any())
            {
                _isWithinCompares.Add("Long", new Func<string, string, string, int>(IsWithinLong));
                _isWithinCompares.Add("Integer", new Func<string, string, string, int>(IsWithinLong));
                _isWithinCompares.Add("Byte", new Func<string, string, string, int>(IsWithinLong));
                _isWithinCompares.Add("Double", new Func<string, string, string, int>(IsWithinDouble));
                _isWithinCompares.Add("Single", new Func<string, string, string, int>(IsWithinDouble));
                _isWithinCompares.Add("Boolean", new Func<string, string, string, int>(IsWithinBoolean));
                _isWithinCompares.Add("Currency", new Func<string, string, string, int>(IsWithinDecimal));
            }

            Func<string, string, string, int> comparer;

            if (!_isWithinCompares.TryGetValue(_typeName, out comparer))
            {
                comparer = IsWithin;
            }
            return comparer;
        }

        private Func<string, string, string, string, int> GetIsStmtComparer()
        {
            if (!_isStmtToIsStmtCompares.Any())
            {
                _isStmtToIsStmtCompares.Add("Long", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtLong));
                _isStmtToIsStmtCompares.Add("Integer", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtLong));
                _isStmtToIsStmtCompares.Add("Byte", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtLong));
                _isStmtToIsStmtCompares.Add("Double", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtDouble));
                _isStmtToIsStmtCompares.Add("Single", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtDouble));
                _isStmtToIsStmtCompares.Add("Boolean", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtBoolean));
                _isStmtToIsStmtCompares.Add("Currency", new Func<string, string, string, string, int>(CompareIsStmtToIsStmtDecimal));
            }

            Func<string, string, string, string, int> comparer;

            if (!_isStmtToIsStmtCompares.TryGetValue(_typeName, out comparer))
            {
                comparer = CompareIsStmtToIsStmt;
            }
            return comparer;
        }

        private Func<string, string, string, string, int> GetRangeToIsStmtComparer()
        {
            if (!_rangeToIsStmtCompares.Any())
            {
                _rangeToIsStmtCompares.Add("Long", new Func<string, string, string, string, int>(CompareRangeToIsStmtLong));
                _rangeToIsStmtCompares.Add("Integer", new Func<string, string, string, string, int>(CompareRangeToIsStmtLong));
                _rangeToIsStmtCompares.Add("Byte", new Func<string, string, string, string, int>(CompareRangeToIsStmtLong));
                _rangeToIsStmtCompares.Add("Double", new Func<string, string, string, string, int>(CompareRangeToIsStmtDouble));
                _rangeToIsStmtCompares.Add("Single", new Func<string, string, string, string, int>(CompareRangeToIsStmtDouble));
                _rangeToIsStmtCompares.Add("Boolean", new Func<string, string, string, string, int>(CompareRangeToIsStmtBoolean));
                _rangeToIsStmtCompares.Add("Currency", new Func<string, string, string, string, int>(CompareRangeToIsStmtDecimal));
            }

            Func<string, string, string, string, int> comparer;

            if (!_rangeToIsStmtCompares.TryGetValue(_typeName, out comparer))
            {
                comparer = CompareIsStmtToIsStmt;
            }
            return comparer;
        }

        private void SetIsParseable()
        {
            if (isLongType)
            {
                long longValue;
                if (_isRange)
                {
                    IsParseable = long.TryParse(_rangeContexts.Key.GetText(), out longValue)
                            && long.TryParse(_rangeContexts.Value.GetText(), out longValue);
                }
                else
                {
                    IsParseable = long.TryParse(GetRangeClauseText(_ctxt), out longValue);
                }
            }
            else if (isDoubleType)
            {
                double dblValue;
                if (_isRange)
                {
                    IsParseable = double.TryParse(_rangeContexts.Key.GetText(), out dblValue)
                            && double.TryParse(_rangeContexts.Value.GetText(), out dblValue);
                }
                else
                {
                    IsParseable = double.TryParse(GetRangeClauseText(_ctxt), out dblValue);
                }
            }
            else if (isBooleanType)
            {
                long longValue;
                if (_isRange)
                {
                    IsParseable = long.TryParse(_rangeContexts.Key.GetText(), out longValue)
                            && long.TryParse(_rangeContexts.Value.GetText(), out longValue);
                }
                else
                {
                    IsParseable = long.TryParse(GetRangeClauseText(_ctxt), out longValue);
                    if (!IsParseable)
                    {
                        IsParseable = (_ctxt.GetText().Equals("True") || _ctxt.GetText().Equals("False"));
                    }
                }
            }
            else if (isDecimalType)
            {
                decimal decimalValue;
                if (_isRange)
                {
                    IsParseable = decimal.TryParse(_rangeContexts.Key.GetText(), out decimalValue)
                            && decimal.TryParse(_rangeContexts.Value.GetText(), out decimalValue);
                }
                else
                {
                    IsParseable = decimal.TryParse(GetRangeClauseText(_ctxt), out decimalValue);
                }
            }
            else
            {
                IsParseable = true;
            }
        }

#region CompareSingleValues
        private int CompareSingleValuesLong(string value1, string value2, string comparisonSymbol)
        {
            return CompareSingleValues(long.Parse(value1), long.Parse(value2), comparisonSymbol);
        }

        private int CompareSingleValuesDouble(string value1, string value2, string comparisonSymbol)
        {
            return CompareSingleValues(double.Parse(value1), double.Parse(value2), comparisonSymbol);
        }

        private int CompareSingleValuesDecimal(string value1, string value2, string comparisonSymbol)
        {
            return CompareSingleValues(decimal.Parse(value1), decimal.Parse(value2), comparisonSymbol);
        }

        private int CompareSingleValuesBoolean(string value1, string value2, string comparisonSymbol)
        {
            return CompareSingleValues(StringToBool(value1), StringToBool(value2), comparisonSymbol);
        }

        private int CompareSingleValues<T>(T priorValue, T candidate, string comparisonSymbol) where T : System.IComparable<T>
        {
            var result = priorValue.CompareTo(candidate);
            if (comparisonSymbol.Equals(EQ))
            {
                return result;
            }
            else if (comparisonSymbol.Equals(NEQ))
            {
                return result != 0 ? 0 : 1;
            }
            else if (comparisonSymbol.Equals(GT))
            {
                return result < 0 ? 0 : result + 1;
            }
            else if (comparisonSymbol.Equals(GTE))
            {
                return result <= 0 ? 0 : result + 1;
            }
            else if (comparisonSymbol.Equals(LT))
            {
                return result > 0 ? 0 : result - 1;
            }
            else if (comparisonSymbol.Equals(LTE))
            {
                return result >= 0 ? 0 : result - 1;
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
        }
#endregion

#region CompareRangeToIsStmt
        private int CompareRangeToIsStmtLong(string priorIsStmtValue, string minVal, string maxVal, string compareSymbol)
        {
            return CompareRangeToIsStmt(long.Parse(priorIsStmtValue), long.Parse(minVal), long.Parse(maxVal), compareSymbol);
        }

        private int CompareRangeToIsStmtDouble(string priorIsStmtValue, string minVal, string maxVal, string compareSymbol)
        {
            return CompareRangeToIsStmt(double.Parse(priorIsStmtValue), double.Parse(minVal), double.Parse(maxVal), compareSymbol);
        }

        private int CompareRangeToIsStmtDecimal(string priorIsStmtValue, string minVal, string maxVal, string compareSymbol)
        {
            return CompareRangeToIsStmt(decimal.Parse(priorIsStmtValue), decimal.Parse(minVal), decimal.Parse(maxVal), compareSymbol);
        }

        private int CompareRangeToIsStmtBoolean(string priorIsStmtValue, string minVal, string maxVal, string compareSymbol)
        {
            return CompareRangeToIsStmt(StringToBool(priorIsStmtValue), StringToBool(minVal), StringToBool(maxVal), compareSymbol);
        }

        private int CompareRangeToIsStmt<T>(T priorIsStmtValue, T minVal, T maxVal, string compareSymbol) where T : System.IComparable<T>
        {
            if (compareSymbol.Equals(GT))
            {
                IsFullyEquivalent = minVal.CompareTo(priorIsStmtValue) > 0;
                IsPartiallyEquivalent = maxVal.CompareTo(priorIsStmtValue) > 0;

                return IsFullyEquivalent || IsPartiallyEquivalent ? 0 : 1;
            }
            else if (compareSymbol.Equals(GTE))
            {
                IsFullyEquivalent = minVal.CompareTo(priorIsStmtValue) >= 0;
                IsPartiallyEquivalent = maxVal.CompareTo(priorIsStmtValue) >= 0;

                return IsFullyEquivalent || IsPartiallyEquivalent ? 0 : 1;
            }
            else if (compareSymbol.Equals(LT))
            {
                IsFullyEquivalent = maxVal.CompareTo(priorIsStmtValue) < 0;
                IsPartiallyEquivalent = minVal.CompareTo(priorIsStmtValue) < 0;

                return IsFullyEquivalent || IsPartiallyEquivalent ? 0 : 1;
            }
            else if (compareSymbol.Equals(LTE))
            {
                IsFullyEquivalent = maxVal.CompareTo(priorIsStmtValue) <= 0;
                IsPartiallyEquivalent = minVal.CompareTo(priorIsStmtValue) <= 0;

                return IsFullyEquivalent || IsPartiallyEquivalent ? 0 : 1;
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
        }
        #endregion

#region CompareSingleValueToIsStmt
        private int CompareSingleValueToIsStmtLong(string value1, string value2, string compareSymbol)
        {
            return CompareSingleValueToIsStmt(long.Parse(value1), long.Parse(value2), compareSymbol);
        }

        private int CompareSingleValueToIsStmtDouble(string value1, string value2, string compareSymbol)
        {
            return CompareSingleValueToIsStmt(double.Parse(value1), double.Parse(value2), compareSymbol);
        }

        private int CompareSingleValueToIsStmtBoolean(string value1, string value2, string compareSymbol)
        {
            return CompareSingleValueToIsStmt(StringToBool(value1), StringToBool(value2), compareSymbol);
        }

        private int CompareSingleValueToIsStmtDecimal(string value1, string value2, string compareSymbol)
        {
            return CompareSingleValueToIsStmt(decimal.Parse(value1), decimal.Parse(value2), compareSymbol);
        }

        private int CompareSingleValueToIsStmt<T>(T isStmtValue, T compareValue, string compareSymbol) where T :  System.IComparable<T>
        {
            var result = isStmtValue.CompareTo(compareValue);
            if (compareSymbol.Equals(EQ))
            {
                return result;
            }
            else if (compareSymbol.Equals(NEQ))
            {
                return result != 0 ? 0 : 1;
            }
            else if (compareSymbol.Equals(GT))
            {
                return result < 0 ? 0 : -1;
            }
            else if (compareSymbol.Equals(GTE))
            {
                return result <= 0 ? 0 : -1;
            }
            else if (compareSymbol.Equals(LT))
            {
                return result > 0 ? 0 : 1;
            }
            else if (compareSymbol.Equals(LTE))
            {
                return result >= 0 ? 0 : 1;
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
        }
        #endregion

#region CompareIsStmtToIsStmt
        private int CompareIsStmtToIsStmtLong(string value1, string value2, string comparisonSymbol1, string comparisonSymbol2)
        {
            return CompareIsStmtToIsStmt(long.Parse(value1), long.Parse(value2), comparisonSymbol1, comparisonSymbol2);
        }

        private int CompareIsStmtToIsStmtDouble(string value1, string value2, string comparisonSymbol1, string comparisonSymbol2)
        {
            return CompareIsStmtToIsStmt(double.Parse(value1), double.Parse(value2), comparisonSymbol1, comparisonSymbol2);
        }

        private int CompareIsStmtToIsStmtDecimal(string value1, string value2, string comparisonSymbol1, string comparisonSymbol2)
        {
            return CompareIsStmtToIsStmt(decimal.Parse(value1), decimal.Parse(value2), comparisonSymbol1, comparisonSymbol2);
        }

        private int CompareIsStmtToIsStmtBoolean(string value1, string value2, string comparisonSymbol1, string comparisonSymbol2)
        {
            return CompareIsStmtToIsStmt(StringToBool(value1), StringToBool(value2), comparisonSymbol1, comparisonSymbol2);
        }

        private int CompareIsStmtToIsStmt<T>(T priorIsStmt, T isStmt, string priorIsStmtCompareSymbol, string isStmtCompareSymbol) where T : System.IComparable<T>
        {
            int returnVal = CompareSingleValueToIsStmt(priorIsStmt, isStmt, priorIsStmtCompareSymbol);
            if (priorIsStmtCompareSymbol.CompareTo(isStmtCompareSymbol) != 0)
            {
                IsPartiallyEquivalent = !(isStmtCompareSymbol.Contains(NEQ) || isStmtCompareSymbol.Contains(EQ));
            }
            else
            {
                IsFullyEquivalent = !(priorIsStmtCompareSymbol.Contains(NEQ) || priorIsStmtCompareSymbol.Contains(EQ));
            }
            return returnVal;
        }
#endregion

#region IsWithin
        private int IsWithinLong(string toCheck, string min, string max)
        {
            return IsWithin(long.Parse(toCheck), long.Parse(min), long.Parse(max));
        }

        private int IsWithinDouble(string toCheck, string min, string max)
        {
            return IsWithin(double.Parse(toCheck), double.Parse(min), double.Parse(max));
        }

        private int IsWithinDecimal(string toCheck, string min, string max)
        {
            return IsWithin(decimal.Parse(toCheck), decimal.Parse(min), decimal.Parse(max));
        }

        private int IsWithinBoolean(string toCheck, string min, string max)
        {
            return IsWithin(StringToBool(toCheck), StringToBool(min), StringToBool(max));
        }

        private int IsWithin<T>(T toCheck, T min, T max) where T : System.IComparable<T>
        {
            return toCheck.CompareTo(min) >= 0 && toCheck.CompareTo(max) <= 0 ? 0 : toCheck.CompareTo(min);
        }
        #endregion

        private string DetermineTheComparisonOperator(VBAParser.RangeClauseContext ctxt)
        {
            _usesIsClause = false;
            var theOperator = EQ;
            //'VBAParser.ComparisonOperatorContext' - The 'Is' case
            var opCtxt = ParserRuleContextHelper.GetChild<VBAParser.ComparisonOperatorContext>(ctxt);
            if (opCtxt != null)
            {
                _usesIsClause = true;
                theOperator = opCtxt.GetText();
            }
            return theOperator;
        }

        private static bool HasChildToken<T>(T ctxt, string token) where T : ParserRuleContext
        {
            var result = false;
            for (int idx = 0; idx < ctxt.ChildCount && !result; idx++)
            {
                if (ctxt.children[idx].GetText().Equals(token))
                {
                    result = true;
                }
            }
            return result;
        }

        private static bool TryGetExprContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : VBAParser.ExpressionContext
        {
            opCtxt = null;
            opCtxt = ParserRuleContextHelper.GetChild<U>(ctxt);
            return opCtxt != null;
        }
    }
}
