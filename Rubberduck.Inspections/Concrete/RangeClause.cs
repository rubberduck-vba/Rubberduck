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
        private readonly string _typeName;
        private bool _usesIsClause;
        private KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext> _rangeContexts;
        private bool _isRange;
        private readonly bool _isSingleVal;
        private string _compareSymbol;
        private bool IsPartiallyEquivalent;
        private bool IsFullyEquivalent;

        public RangeClause(VBAParser.RangeClauseContext ctxt, IdentifierReference theRef)
        {
            _ctxt = ctxt;
            _theRef = theRef;
            _typeName = theRef.Declaration.AsTypeName;
            _compareSymbol = DetermineTheComparisonOperator(ctxt);
            //_isRange = 
            TryGetRangeContext(ctxt, out _rangeContexts);
            _usesIsClause = UsesIsExpression(ctxt);
            _isRange = UsesRangeExpression(ctxt);
            _isSingleVal = !_isRange;

            //TODO: Below are type validations that belong in early checks...maybe before calling this constructor
            var result = GetRangeClauseText(ctxt);
            if (HandleAsLong(theRef.Declaration.AsTypeName) && _isSingleVal)
            {
                long longValue;
                var test = long.TryParse(result, out longValue);
                if (!test)
                {
                    _typeName = "String";
                }
            }
            if (HandleAsDouble(theRef.Declaration.AsTypeName) && _isSingleVal)
            {
                double dblValue;
                var test = double.TryParse(result, out dblValue);
                if (!test)
                {
                    _typeName = "String";
                }
            }
        }

        public string ClauseTypeName => _typeName;
        public bool IsSingleVal => _isSingleVal;
        public bool UsesIsClause => _usesIsClause;
        public bool IsRange => _isRange;
        public string CompareSymbol => _compareSymbol;
        public string TypeName => _typeName;
        private IdentifierReference _theRef;

        private bool isLongType => HandleAsLong(_typeName);
        private bool isDoubleType => HandleAsDouble(_typeName);

        public string ValueAsString => _typeName.Equals("String") ? _ctxt.GetText() : GetRangeClauseText(_ctxt);
        public string ValueMinAsString => _isRange ? _rangeContexts.Key.GetText() : ValueAsString;
        public string ValueMaxAsString => _isRange ? _rangeContexts.Value.GetText() : ValueAsString;

        private static bool UsesIsExpression(VBAParser.RangeClauseContext ctxt)
        {
            var usesIsClause = false;
            for (int idx = 0; idx < ctxt.ChildCount && !usesIsClause; idx++)
            {
                if (ctxt.children[idx].GetText().Equals("Is"))
                {
                    usesIsClause = true;
                }
            }
            return usesIsClause;
        }

        private static bool UsesRangeExpression(VBAParser.RangeClauseContext ctxt)
        {
            var isRange = false;
            for (int idx = 0; idx < ctxt.ChildCount && !isRange; idx++)
            {
                if (ctxt.children[idx].GetText().Equals("To"))
                {
                    isRange = true;
                }
            }
            return isRange;
        }

        private string GetRangeClauseText(VBAParser.RangeClauseContext ctxt)
        {
            var relationalOpCtxt = ParserRuleContextHelper.GetChild<VBAParser.RelationalOpContext>(ctxt);
            if(relationalOpCtxt != null)
            {
                var lExprCtxtIndex = -1;
                var literalExprCtxtIndex = -1;
                var lExprCtxt = ParserRuleContextHelper.GetChild<VBAParser.LExprContext>(relationalOpCtxt);
                if(lExprCtxt != null)
                {
                    var theValueCtxt = ParserRuleContextHelper.GetChild<VBAParser.LiteralExprContext>(relationalOpCtxt);
                    if(theValueCtxt == null)
                    {
                        return ""; //TODO: handle this better...is it an error to be detected earlier?
                    }
                    for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
                    {
                        if (relationalOpCtxt.children[idx] is VBAParser.LExprContext)
                        {
                            lExprCtxtIndex = idx;
                        }
                        else if (relationalOpCtxt.children[idx] is VBAParser.LiteralExprContext)
                        {
                            literalExprCtxtIndex = idx;
                        }
                    }
                    for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
                    {

                        var content = relationalOpCtxt.children[idx].GetText();
                        if (ComparisonOperators.Contains(content))
                        {
                            if (lExprCtxtIndex < literalExprCtxtIndex)
                            {
                                _compareSymbol = content;
                            }
                            else
                            {
                                _compareSymbol = OperatorInversions[content];
                            }
                        }
                    }
                    if (lExprCtxt.GetText().Equals(_theRef.IdentifierName))
                    {
                        _usesIsClause = true;
                        return theValueCtxt.GetText();
                    }
                }
            }

            var negativeCtxt = ParserRuleContextHelper.GetChild<VBAParser.UnaryMinusOpContext>(ctxt);
            if (negativeCtxt != null)
            {
                var theValueCtxt = ParserRuleContextHelper.GetChild<VBAParser.LiteralExprContext>(negativeCtxt);
                return theValueCtxt != null ? negativeCtxt.GetText() + theValueCtxt.GetText() : string.Empty;
            }
            else
            {
                var theValueCtxt = ParserRuleContextHelper.GetChild<VBAParser.LiteralExprContext>(ctxt);
                return theValueCtxt != null ? theValueCtxt.GetText() : string.Empty;
            }
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
                if (!UsesIsClause && !prior.UsesIsClause)
                {
                    var result = 1;
                    if (isLongType)
                    {
                        result = CompareSingleValues(long.Parse(prior.ValueAsString), long.Parse(ValueAsString), EQ);
                    }
                    else if (isDoubleType)
                    {
                        result = CompareSingleValues(double.Parse(prior.ValueAsString), double.Parse(ValueAsString), EQ);
                    }
                    else
                    {
                        result = CompareSingleValues(prior.ValueAsString, ValueAsString, EQ);
                    }
                    IsFullyEquivalent = result == 0;
                    return result;
                }
                else if (!UsesIsClause && prior.UsesIsClause)
                {
                    var result = 1;
                    if (isLongType)
                    {
                        result = CompareSingleValues(long.Parse(prior.ValueAsString), long.Parse(ValueAsString), prior.CompareSymbol);
                    }
                    else if (isDoubleType)
                    {
                        result = CompareSingleValues(double.Parse(prior.ValueAsString), double.Parse(ValueAsString), prior.CompareSymbol);
                    }
                    IsFullyEquivalent = (result == 0);
                    return result;
                }
                else if (UsesIsClause && !prior.UsesIsClause)
                {
                    var result = 1;
                    //TODO: consider explicit re-statment of comparison symbol rather than swapping params
                    if (isLongType)
                    {
                        result = CompareSingleValues(long.Parse(ValueAsString), long.Parse(prior.ValueAsString), CompareSymbol);
                    }
                    if (isDoubleType)
                    {
                        result = CompareSingleValues(double.Parse(ValueAsString), double.Parse(prior.ValueAsString), CompareSymbol);
                    }
                    else
                    {
                        result = CompareSingleValues(prior.ValueAsString, ValueAsString, CompareSymbol);
                    }
                    IsPartiallyEquivalent = (result == 0);
                    return result;
                }
                else if (UsesIsClause && prior.UsesIsClause)
                {
                    if (isLongType)
                    {
                        return CompareIsStmtToIsStmt(long.Parse(prior.ValueAsString), long.Parse(ValueAsString), prior.CompareSymbol, CompareSymbol);
                    }
                    else if (isDoubleType)
                    {
                        return CompareIsStmtToIsStmt(double.Parse(prior.ValueAsString), double.Parse(ValueAsString), prior.CompareSymbol, CompareSymbol);
                    }
                    else
                    {
                        return CompareIsStmtToIsStmt(prior.ValueAsString, ValueAsString, prior.CompareSymbol, CompareSymbol);
                    }
                }
                Debug.Assert(true, "Unanticipated code path");
                return 0;
            }
            else if (IsSingleVal && prior.IsRange)
            {
                if (!UsesIsClause)
                {
                    var result = 1;
                    if (isLongType)
                    {
                        result = IsWithin(long.Parse(ValueAsString), long.Parse(prior.ValueMinAsString), long.Parse(prior.ValueMaxAsString));
                    }
                    else if (isDoubleType)
                    {
                        result = IsWithin(double.Parse(ValueAsString), double.Parse(prior.ValueMinAsString), double.Parse(prior.ValueMaxAsString));
                    }
                    else
                    {
                        result = IsWithin(ValueAsString, prior.ValueMinAsString, prior.ValueMaxAsString);
                    }
                    IsFullyEquivalent = result == 0;
                    return result;
                }
                else
                {
                    int resultStartVal = 1;
                    int resultEndVal = 1;
                    // this Is > 8 prior 3 to 10
                    if (isLongType)
                    {
                        resultStartVal = CompareSingleValueToIsStmt(long.Parse(ValueAsString), long.Parse(prior.ValueMinAsString), CompareSymbol);
                        resultEndVal = CompareSingleValueToIsStmt(long.Parse(ValueAsString), long.Parse(prior.ValueMaxAsString), CompareSymbol);
                    }
                    if (isDoubleType)
                    {
                        resultStartVal = CompareSingleValueToIsStmt(double.Parse(ValueAsString), double.Parse(prior.ValueMinAsString), CompareSymbol);
                        resultEndVal = CompareSingleValueToIsStmt(double.Parse(ValueAsString), double.Parse(prior.ValueMaxAsString), CompareSymbol);
                    }
                    return resultStartVal == 0 || resultEndVal == 0 ? 0 : 1;
                }
            }
            else if (IsRange && prior.IsSingleVal)
            {
                if (!prior.UsesIsClause)
                {
                    //var result = IsWithin(prior.ValueMin, ValueMin, ValueMax) * -1;
                    var result = 1;
                    if (isLongType)
                    {
                        result = IsWithin(long.Parse(prior.ValueAsString), long.Parse(ValueMinAsString), long.Parse(ValueMaxAsString));
                    }
                    else if (isDoubleType)
                    {
                        result = IsWithin(double.Parse(prior.ValueAsString), double.Parse(ValueMinAsString), double.Parse(ValueMaxAsString));
                    }
                    else
                    {
                        result = IsWithin(prior.ValueAsString, ValueMinAsString, ValueMaxAsString);
                    }
                    IsPartiallyEquivalent = result == 0;
                    return result;
                }
                else
                {
                    if (isLongType)
                    {
                        return CompareRangeToIsStmt(long.Parse(prior.ValueMinAsString), long.Parse(ValueMinAsString), long.Parse(ValueMaxAsString), prior.CompareSymbol);
                    }
                    else if (isDoubleType)
                    {
                        return CompareRangeToIsStmt(double.Parse(prior.ValueMinAsString), double.Parse(ValueMinAsString), double.Parse(ValueMaxAsString), prior.CompareSymbol);
                    }
                    else
                    {
                        return CompareRangeToIsStmt(prior.ValueMinAsString, ValueMinAsString, ValueMaxAsString, prior.CompareSymbol);
                    }
                }
            }
            else if (IsRange && prior.IsRange)
            {
                if (isLongType)
                {
                    if (IsWithin(long.Parse(ValueMinAsString), long.Parse(prior.ValueMinAsString), long.Parse(prior.ValueMaxAsString)) == 0
                         && IsWithin(long.Parse(ValueMaxAsString), long.Parse(prior.ValueMinAsString), long.Parse(prior.ValueMaxAsString)) == 0)
                    {
                        IsFullyEquivalent = true;
                        return 0;
                    }
                    else
                    {
                        IsPartiallyEquivalent = IsWithin(long.Parse(ValueMinAsString), long.Parse(prior.ValueMinAsString), long.Parse(prior.ValueMaxAsString)) == 0
                            || IsWithin(long.Parse(ValueMaxAsString), long.Parse(prior.ValueMinAsString), long.Parse(prior.ValueMaxAsString)) == 0;

                        return IsPartiallyEquivalent ? 0 : ValueMaxAsString.CompareTo(prior.ValueMaxAsString);
                    }
                }
                else if (isDoubleType)
                {
                    if (IsWithin(double.Parse(ValueMinAsString), double.Parse(prior.ValueMinAsString), double.Parse(prior.ValueMaxAsString)) == 0
                         && IsWithin(double.Parse(ValueMaxAsString), double.Parse(prior.ValueMinAsString), double.Parse(prior.ValueMaxAsString)) == 0)
                    {
                        IsFullyEquivalent = true;
                        return 0;
                    }
                    else
                    {
                        IsPartiallyEquivalent = IsWithin(double.Parse(ValueMinAsString), double.Parse(prior.ValueMinAsString), double.Parse(prior.ValueMaxAsString)) == 0
                            || IsWithin(double.Parse(ValueMaxAsString), double.Parse(prior.ValueMinAsString), double.Parse(prior.ValueMaxAsString)) == 0;

                        return IsPartiallyEquivalent ? 0 : ValueMaxAsString.CompareTo(prior.ValueMaxAsString);
                    }
                }
                else
                {
                    if (IsWithin(ValueMinAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0
                         && IsWithin(ValueMaxAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0)
                    {
                        IsFullyEquivalent = true;
                        return 0;
                    }
                    else
                    {
                        IsPartiallyEquivalent = IsWithin(ValueMinAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0
                            || IsWithin(ValueMaxAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0;

                        return IsPartiallyEquivalent ? 0 : ValueMaxAsString.CompareTo(prior.ValueMaxAsString);
                    }
                }
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
        }

        private int CompareSingleValues<T>( T priorRgValue, T candidate, string comparisonSymbol) where T : System.IComparable<T>
        {
            var result =  priorRgValue.CompareTo(candidate);
            if (comparisonSymbol.Equals(GT))
            {
                return result < 0 ? 0 : result + 1;
            }
            if (comparisonSymbol.Equals(GTE))
            {
                return result <= 0 ? 0 : result + 1;
            }
            if (comparisonSymbol.Equals(LT))
            {
                return result > 0 ? 0 : result - 1;
            }
            if (comparisonSymbol.Equals(LTE))
            {
                return result >= 0 ? 0 : result - 1;
            }
            if (comparisonSymbol.Equals(EQ))
            {
                return result;
            }
            if (comparisonSymbol.Equals(NEQ))
            {
                return result != 0 ? 0 : 1;
            }
            else
            {
                return result;
            }

        }

        private int CompareRangeToIsStmt<T>(T priorIsStmtValue, T minVal, T maxVal, string compareSymbol) where T : System.IComparable<T>
        {
            if (compareSymbol.Equals(GT))
            {
                if (minVal.CompareTo(priorIsStmtValue) > 0)
                {
                    return 0;
                }
                else if (maxVal.CompareTo(priorIsStmtValue) > 0)
                {
                    return 0;
                }
                return -1;
            }
            else if (compareSymbol.Equals(GTE))
            {
                if (minVal.CompareTo(priorIsStmtValue) >= 0)
                {
                    return 0;
                }
                else if (maxVal.CompareTo(priorIsStmtValue) >= 0)
                {
                    return 0;
                }
                return -1;
            }
            else if (compareSymbol.Equals(LT))
            {
                if (maxVal.CompareTo(priorIsStmtValue) < 0)
                {
                    return 0; //Full
                }
                else if (minVal.CompareTo(priorIsStmtValue) < 0)
                {
                    return 0; //Partial
                }
                return 1;
            }
            else if (compareSymbol.Equals(LTE))
            {
                if (maxVal.CompareTo(priorIsStmtValue) <= 0)
                {
                    return 0; //Full
                }
                else if (minVal.CompareTo(priorIsStmtValue) <= 0)
                {
                    return 0; //Partial
                }
                return 1;
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
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

        private int CompareIsStmtToIsStmt<T>(T isStmtValuePreceding, T isStmtValueSecond, string compareSymbolPrecedingIsStmt, string compareSymbolSecond) where T : System.IComparable<T>
        {
            int returnVal = CompareSingleValueToIsStmt(isStmtValuePreceding, isStmtValueSecond, compareSymbolPrecedingIsStmt);
            if (compareSymbolPrecedingIsStmt.CompareTo(compareSymbolSecond) != 0)
            {
                IsPartiallyEquivalent = !(compareSymbolSecond.Contains(NEQ) || compareSymbolSecond.Contains(EQ));
            }
            else
            {
                IsFullyEquivalent = !(compareSymbolPrecedingIsStmt.Contains(NEQ) || compareSymbolPrecedingIsStmt.Contains(EQ));
            }
            return returnVal;
        }

        private int IsWithin<T>(T toCheck, T min, T max) where T : System.IComparable<T>
        {
            return toCheck.CompareTo(min) >= 0 && toCheck.CompareTo(max) <= 0 ? 0 : toCheck.CompareTo(min);
        }

        private string DetermineTheComparisonOperator(VBAParser.RangeClauseContext ctxt)
        {
            _usesIsClause = false;
            var theOperator = EQ;
            //'VBAParser.ComparisonOperatorContext' is the The 'Is' case
            var opCtxt = ParserRuleContextHelper.GetChild<VBAParser.ComparisonOperatorContext>(ctxt);
            if (opCtxt != null)
            {
                _usesIsClause = true;
                theOperator = opCtxt.GetText();
            }
            return theOperator;
        }

        private bool TryGetRangeContext(VBAParser.RangeClauseContext ctxt, out KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext> rangeContexts)
        {
            rangeContexts = new KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext>
                (ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(ctxt),
                ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(ctxt));

            return rangeContexts.Key != null && rangeContexts.Value != null;
        }

        private static bool HandleAsLong(String typeName)
        {
            string[] types = { "Integer", "Long", "Single", "Byte" };
            return types.Contains(typeName);
        }

        private static bool HandleAsDouble(String typeName)
        {
            string[] types = { "Double", "Double(negative)", "Double(positive)", "Currency" };
            return types.Contains(typeName);
        }

        private static bool HandleAsBoolean(String typeName)
        {
            string[] types = { "Boolean" };
            return types.Contains(typeName);
        }

        private static string[] ComparisonOperators = { EQ, NEQ, LT, LTE, GT, GTE };
        private static Dictionary<string, string> OperatorInversions = new Dictionary<string, string>()
        {
            { EQ,EQ },
            {NEQ,NEQ },
            {LT,GT },
            {LTE,GTE },
            {GT,LT },
            {GTE,LTE }
        };
    }
}
