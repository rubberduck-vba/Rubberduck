using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
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
        private bool _isParseable;
        private bool IsPartiallyEquivalent;
        private bool IsFullyEquivalent;

        private static string[] _comparisonOperators = { EQ, NEQ, LT, LTE, GT, GTE };
        private static Dictionary<string, string> _operatorInversions = new Dictionary<string, string>()
        {
            { EQ,EQ },
            { NEQ,NEQ },
            { LT,GT },
            { LTE,GTE },
            { GT,LT },
            { GTE,LTE }
        };

        public RangeClause(VBAParser.RangeClauseContext ctxt, IdentifierReference theRef)
        {
            _ctxt = ctxt;
            _theRef = theRef;
            _typeName = theRef.Declaration.AsTypeName;
            _compareSymbol = DetermineTheComparisonOperator(ctxt);
            _usesIsClause = HasChildToken(ctxt, Tokens.Is);
            _isSingleVal = true;
            _isParseable = true;
            _isRange = HasChildToken(ctxt, Tokens.To);
            if (_isRange)
            {
                _rangeContexts = new KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext>
                    (ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(ctxt),
                    ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(ctxt));
                _isParseable = _rangeContexts.Key != null && _rangeContexts.Value != null;
                _isSingleVal = false;
            }

            if (HandleAsLong(theRef.Declaration.AsTypeName))// && _isSingleVal)
            {
                long longValue;
                if (_isRange)
                {
                    _isParseable = long.TryParse(_rangeContexts.Key.GetText(), out longValue)
                            && long.TryParse(_rangeContexts.Value.GetText(), out longValue);
                }
                else
                {
                    _isParseable = long.TryParse(GetRangeClauseText(ctxt), out longValue);
                }
            }
            else if (HandleAsDouble(theRef.Declaration.AsTypeName))// && _isSingleVal)
            {
                double dblValue;
                if (_isRange)
                {
                    _isParseable = double.TryParse(_rangeContexts.Key.GetText(), out dblValue)
                            && double.TryParse(_rangeContexts.Value.GetText(), out dblValue);
                }
                else
                {
                    _isParseable = double.TryParse(GetRangeClauseText(ctxt), out dblValue);
                }
            }
            else if (HandleAsDecimal(theRef.Declaration.AsTypeName))// && _isSingleVal)
            {
                decimal decimalValue;
                if (_isRange)
                {
                    _isParseable = decimal.TryParse(_rangeContexts.Key.GetText(), out decimalValue)
                            && decimal.TryParse(_rangeContexts.Value.GetText(), out decimalValue);
                }
                else
                {
                    _isParseable = decimal.TryParse(GetRangeClauseText(ctxt), out decimalValue);
                }
            }
            else if (HandleAsBoolean(theRef.Declaration.AsTypeName))// && _isSingleVal)
            {
                long longValue;
                if (_isRange)
                {
                    _isParseable = long.TryParse(_rangeContexts.Key.GetText(), out longValue)
                            && long.TryParse(_rangeContexts.Value.GetText(), out longValue);
                }
                else
                {

                    _isParseable = long.TryParse(GetRangeClauseText(ctxt), out longValue);
                    if (!_isParseable)
                    {
                        _isParseable = (ctxt.GetText().Equals("True") || ctxt.GetText().Equals("False"));
                    }
                }
            }
        }

        public string ClauseTypeName => _typeName;
        public bool IsSingleVal => _isSingleVal;
        public bool UsesIsClause => _usesIsClause;
        public bool IsRange => _isRange;
        public string CompareSymbol => _compareSymbol;
        public string TypeName => _typeName;
        public bool IsParseable => _isParseable;
        private IdentifierReference _theRef;

        private bool isLongType => HandleAsLong(_typeName);
        private bool isDoubleType => HandleAsDouble(_typeName);
        private bool isBooleanType => HandleAsBoolean(_typeName);
        private bool isDecimalType => HandleAsDecimal(_typeName);

        public string ValueAsString => _typeName.Equals("String") ? _ctxt.GetText() : GetRangeClauseText(_ctxt);
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

            //TODO: Figure out how to use 'GetToken' - and see if the comparison symbols can be acquired there.
            // var test = (ParserRuleContext)relationalOpCtxt.GetToken(VBAParser.GT, 0);
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
                else if (_comparisonOperators.Contains(text))
                {
                    _compareSymbol = text;
                }
            }

            if (lExprCtxtIndex > literalExprCtxtIndex)
            {
                _compareSymbol = _operatorInversions[_compareSymbol];
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
            int intVal = 0;
            if (strValue.Equals("True") || strValue.Equals("False"))
            {
                intVal = strValue.Equals("True") ? 1 : 0;
            }
            else
            {
                intVal = int.Parse(strValue);
            }
            return intVal != 0;
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
                    else if (isDecimalType)
                    {
                        result = CompareSingleValues(decimal.Parse(prior.ValueAsString), decimal.Parse(ValueAsString), EQ);
                    }
                    else if (isBooleanType)
                    {
                        result = CompareSingleValues(StringToBool(prior.ValueAsString), StringToBool(ValueAsString), EQ);
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
                    else if (isDecimalType)
                    {
                        result = CompareSingleValues(decimal.Parse(prior.ValueAsString), decimal.Parse(ValueAsString), prior.CompareSymbol);
                    }
                    else if (isBooleanType)
                    {
                        result = CompareSingleValues(StringToBool(prior.ValueAsString), StringToBool(ValueAsString), prior.CompareSymbol);
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
                    else if (isDoubleType)
                    {
                        result = CompareSingleValues(double.Parse(ValueAsString), double.Parse(prior.ValueAsString), CompareSymbol);
                    }
                    else if (isDecimalType)
                    {
                        result = CompareSingleValues(decimal.Parse(ValueAsString), decimal.Parse(prior.ValueAsString), CompareSymbol);
                    }
                    else if (isBooleanType)
                    {
                        result = CompareSingleValues(StringToBool(ValueAsString), StringToBool(prior.ValueAsString), CompareSymbol);
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
                    else if (isDecimalType)
                    {
                        return CompareIsStmtToIsStmt(decimal.Parse(prior.ValueAsString), decimal.Parse(ValueAsString), prior.CompareSymbol, CompareSymbol);
                    }
                    else if (isBooleanType)
                    {
                        return CompareIsStmtToIsStmt(StringToBool(prior.ValueAsString), StringToBool(ValueAsString), prior.CompareSymbol, CompareSymbol);
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
                    else if (isDecimalType)
                    {
                        result = IsWithin(decimal.Parse(ValueAsString), decimal.Parse(prior.ValueMinAsString), decimal.Parse(prior.ValueMaxAsString));
                    }
                    else if (isBooleanType)
                    {
                        //result = IsWithin(StringToBool(ValueAsString), long.Parse(prior.ValueMinAsString) != 0, long.Parse(prior.ValueMaxAsString) != 0);
                        result = IsWithin(StringToBool(ValueAsString), StringToBool(prior.ValueMinAsString), StringToBool(prior.ValueMaxAsString));
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
                    else if (isDoubleType)
                    {
                        resultStartVal = CompareSingleValueToIsStmt(double.Parse(ValueAsString), double.Parse(prior.ValueMinAsString), CompareSymbol);
                        resultEndVal = CompareSingleValueToIsStmt(double.Parse(ValueAsString), double.Parse(prior.ValueMaxAsString), CompareSymbol);
                    }
                    else if (isDecimalType)
                    {
                        resultStartVal = CompareSingleValueToIsStmt(decimal.Parse(ValueAsString), decimal.Parse(prior.ValueMinAsString), CompareSymbol);
                        resultEndVal = CompareSingleValueToIsStmt(decimal.Parse(ValueAsString), decimal.Parse(prior.ValueMaxAsString), CompareSymbol);
                    }
                    else if (isBooleanType)
                    {
                        resultStartVal = CompareSingleValueToIsStmt(StringToBool(ValueAsString), StringToBool(prior.ValueMinAsString), CompareSymbol);
                        resultEndVal = CompareSingleValueToIsStmt(StringToBool(ValueAsString), StringToBool(prior.ValueMaxAsString), CompareSymbol);
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
                    else if (isDecimalType)
                    {
                        result = IsWithin(decimal.Parse(prior.ValueAsString), decimal.Parse(ValueMinAsString), decimal.Parse(ValueMaxAsString));
                    }
                    else if (isBooleanType)
                    {
                        result = IsWithin(StringToBool(prior.ValueAsString), StringToBool(ValueMinAsString), StringToBool(ValueMaxAsString));
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
                    else if (isDecimalType)
                    {
                        return CompareRangeToIsStmt(decimal.Parse(prior.ValueMinAsString), decimal.Parse(ValueMinAsString), decimal.Parse(ValueMaxAsString), prior.CompareSymbol);
                    }
                    else if (isBooleanType)
                    {
                        //return CompareRangeToIsStmt(decimal.Parse(prior.ValueMinAsString), decimal.Parse(ValueMinAsString), decimal.Parse(ValueMaxAsString), prior.CompareSymbol);
                        return CompareRangeToIsStmt(StringToBool(prior.ValueMinAsString), StringToBool(ValueMinAsString), StringToBool(ValueMaxAsString), prior.CompareSymbol);
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
                else if (isDecimalType)
                {
                    if (IsWithin(decimal.Parse(ValueMinAsString), decimal.Parse(prior.ValueMinAsString), decimal.Parse(prior.ValueMaxAsString)) == 0
                         && IsWithin(decimal.Parse(ValueMaxAsString), decimal.Parse(prior.ValueMinAsString), decimal.Parse(prior.ValueMaxAsString)) == 0)
                    {
                        IsFullyEquivalent = true;
                        return 0;
                    }
                    else
                    {
                        IsPartiallyEquivalent = IsWithin(decimal.Parse(ValueMinAsString), decimal.Parse(prior.ValueMinAsString), decimal.Parse(prior.ValueMaxAsString)) == 0
                            || IsWithin(decimal.Parse(ValueMaxAsString), decimal.Parse(prior.ValueMinAsString), decimal.Parse(prior.ValueMaxAsString)) == 0;

                        return IsPartiallyEquivalent ? 0 : ValueMaxAsString.CompareTo(prior.ValueMaxAsString);
                    }
                }
                else if (isBooleanType)
                {
                    if (IsWithin(StringToBool(ValueMinAsString), StringToBool(prior.ValueMinAsString), StringToBool(prior.ValueMaxAsString)) == 0
                         && IsWithin(StringToBool(ValueMaxAsString), StringToBool(prior.ValueMinAsString), StringToBool(prior.ValueMaxAsString)) == 0)
                    {
                        IsFullyEquivalent = true;
                        return 0;
                    }
                    else
                    {
                        IsPartiallyEquivalent = IsWithin(bool.Parse(ValueMinAsString), bool.Parse(prior.ValueMinAsString), bool.Parse(prior.ValueMaxAsString)) == 0
                            || IsWithin(bool.Parse(ValueMaxAsString), bool.Parse(prior.ValueMinAsString), bool.Parse(prior.ValueMaxAsString)) == 0;

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

        private static bool HandleAsLong(String typeName)
        {
            string[] types = { "Integer", "Long", "Single", "Byte" };
            return types.Contains(typeName);
        }

        private static bool HandleAsDouble(String typeName)
        {
            string[] types = { "Double", "Double(negative)", "Double(positive)" };
            return types.Contains(typeName);
        }

        private static bool HandleAsDecimal(String typeName)
        {
            string[] types = { "Currency" };
            return types.Contains(typeName);
        }

        private static bool HandleAsBoolean(String typeName)
        {
            string[] types = { "Boolean" };
            return types.Contains(typeName);
        }
    }
}
