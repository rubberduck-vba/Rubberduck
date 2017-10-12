using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Inspections.Concrete
{
    public struct RangeClauseEvaluationResults
    {
        public bool IsReachable;
        public bool IsParseable;
        public bool IsStringLiteral;
        public bool MatchesSelectCaseTypeName;
        public bool IsFullyEquivalent;
        public bool IsPartiallyEquivalent;

        public RangeClauseEvaluationResults(IRangeClause rangeClause)
        {
            IsReachable = true;
            IsParseable = true;
            IsStringLiteral = true;
            MatchesSelectCaseTypeName = true;
            IsFullyEquivalent = false;
            IsPartiallyEquivalent = false;
        }
    }

    public class RangeClause : IRangeClause
    {
        private const string EQ = "=";
        private const string NEQ = "<>";
        private const string LT = "<";
        private const string LTE = "<=";
        private const string GT = ">";
        private const string GTE = ">=";

        private readonly VBAParser.RangeClauseContext _ctxt;
        private readonly string _typeName;
        RubberduckParserState _state;
        private bool _usesIsClause;
        private KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext> _rangeContexts;
        private bool _isRange;
        private bool _RangeValuesAreHighToLow;
        private readonly bool _isSingleVal;
        private string _compareSymbol;
        public bool IsPartiallyEquivalent { get; set; }
        public bool IsFullyEquivalent { get; set; }
        public bool IsStringLiteral { get; set; }

        private Func<string, string, string, int> SingleValueClauseCompare;
        private Func<string, string, string, int> IsWithinClauseCompare;
        private Func<string, string, string, int> SingleValueToIsStmtClauseCompare;
        private Func<string, string, string, string, int> IsStmtToIsStmtClauseCompare;
        private Func<string, string, string, string, int> RangeToIsStmtClauseCompare;
        private Func<string, string, int> StandardValueToValueCompare;

        private RangeClauseEvaluationResults _evalResults;


        private static string[] LongComparisonTypes = { "Integer", "Long", "Byte" };
        private static string[] DoubleComparisonTypes = { "Double", "Single" };
        private static string[] CurrencyComparisonTypes = { "Currency" };
        private static string[] BooleanComparisonTypes = { "Boolean" };

        private static Dictionary<string, string> _comparisonOperatorsAndInversions = new Dictionary<string, string>()
        {
            { EQ,NEQ },
            { NEQ,EQ },
            { LT,GT },
            { LTE,GTE },
            { GT,LT },
            { GTE,LTE }
        };

        private static Dictionary<string, string> _extendeCcomparisonOperatorInversions = new Dictionary<string, string>()
        {
            { LT,GTE },
            { LTE,GT },
            { GT,LTE },
            { GTE,LT }
        };

        public RangeClause(RubberduckParserState state, VBAParser.RangeClauseContext ctxt, IdentifierReference theRef, string typeName)
        {
            _state = state;
            _ctxt = ctxt;
            _theRef = theRef;
            _typeName = typeName;
            SetCompareDelegates();
            _compareSymbol = DetermineTheComparisonOperator(ctxt);
            _usesIsClause = HasChildToken(ctxt, Tokens.Is);
            _isRange = HasChildToken(ctxt, Tokens.To);
            _isSingleVal = !_isRange;
            _evalResults = new RangeClauseEvaluationResults
            {
                IsParseable = true,
                IsFullyEquivalent = false,
                IsPartiallyEquivalent = false,
                MatchesSelectCaseTypeName = true,
                IsReachable = true,
                IsStringLiteral = false
            };

            if (_isRange)
            {
                _rangeContexts = new KeyValuePair<VBAParser.SelectStartValueContext, VBAParser.SelectEndValueContext>
                    (ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(_ctxt),
                    ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(_ctxt));
                IsParseable = _rangeContexts.Key != null && _rangeContexts.Value != null;
                _RangeValuesAreHighToLow = StandardValueToValueCompare(GetText(_rangeContexts.Key), GetText(_rangeContexts.Value)) > 0;
            }
            IsStringLiteral = _ctxt.GetText().StartsWith("\"") && _ctxt.GetText().EndsWith("\"");
            _evalResults.IsStringLiteral = IsStringLiteral;

            SetIsParseable();

            HasUnreachableCaseElse = false;
        }

        public bool IsSingleVal => _isSingleVal;
        public bool UsesIsClause => _usesIsClause;
        public bool IsRange => _isRange;
        public string CompareSymbol => _compareSymbol;
        private string SelectCaseTypeName => _typeName;
        public bool IsParseable { get; set; }
        public bool MatchesSelectCaseType
        {
            get
            {
                return _evalResults.MatchesSelectCaseTypeName;
            } 
        }
        public bool HasUnreachableCaseElse { get; set; }

        private bool IsComparisonOperator(string opCandidate) { return _comparisonOperatorsAndInversions.Keys.Contains(opCandidate); }
        private string GetOperatorInverseStrict(string theOperator)
        {
            return IsComparisonOperator(theOperator) ? _comparisonOperatorsAndInversions[theOperator] : theOperator;
        }
        private string GetOperatorInverseExtendedSet(string theOperator)
        {
            return IsComparisonOperator(theOperator) ? _extendeCcomparisonOperatorInversions[theOperator] : theOperator;
        }

        private IdentifierReference _theRef;

        private bool isLongType => LongComparisonTypes.Contains(SelectCaseTypeName);
        private bool isDoubleType => DoubleComparisonTypes.Contains(SelectCaseTypeName);
        private bool isBooleanType => BooleanComparisonTypes.Contains(SelectCaseTypeName);
        private bool isDecimalType => CurrencyComparisonTypes.Contains(SelectCaseTypeName);
        private bool isStringType => !(isLongType || isDoubleType || isBooleanType || isDecimalType);

        public string ValueAsString => GetRangeClauseText(_ctxt);
        public string ValueMinAsString => _isRange ? _RangeValuesAreHighToLow ? GetText( _rangeContexts.Value) : GetText(_rangeContexts.Key) : ValueAsString;
        public string ValueMaxAsString => _isRange ? _RangeValuesAreHighToLow ? GetText(_rangeContexts.Key) : GetText(_rangeContexts.Value) : ValueAsString;


        public bool IsReachable(object obj)
        {
            var prior = obj as IRangeClause;
            if (prior == null)
            {
                throw new InvalidCastException("Unable to cast parameter 'obj' to IRangeClass");
            }

            if (!_evalResults.MatchesSelectCaseTypeName)
            {
                return false;
            }

            return CompareTo(prior) != 0;
        }

        private void SetCompareDelegates()
        {
            if (isLongType)
            {
                SingleValueClauseCompare = CompareSingleValuesLong;
                IsWithinClauseCompare = CompareIsWithinLong;
                SingleValueToIsStmtClauseCompare = CompareSingleValueToIsStmtLong;
                IsStmtToIsStmtClauseCompare = CompareIsStmtToIsStmtLong;
                RangeToIsStmtClauseCompare = CompareRangeToIsStmtLong;
                StandardValueToValueCompare = SimpleCompareLong;
            }
            else if (isDecimalType)
            {
                SingleValueClauseCompare = CompareSingleValuesDecimal;
                IsWithinClauseCompare = CompareIsWithinDecimal;
                SingleValueToIsStmtClauseCompare = CompareSingleValueToIsStmtDecimal;
                IsStmtToIsStmtClauseCompare = CompareIsStmtToIsStmtDecimal;
                RangeToIsStmtClauseCompare = CompareRangeToIsStmtDecimal;
                StandardValueToValueCompare = SimpleCompareDecimal;
            }
            else if (isDoubleType)
            {
                SingleValueClauseCompare = CompareSingleValuesDouble;
                IsWithinClauseCompare = CompareIsWithinDouble;
                SingleValueToIsStmtClauseCompare = CompareSingleValueToIsStmtDouble;
                IsStmtToIsStmtClauseCompare = CompareIsStmtToIsStmtDouble;
                RangeToIsStmtClauseCompare = CompareRangeToIsStmtDouble;
                StandardValueToValueCompare = SimpleCompareDouble;
            }
            else if (isBooleanType)
            {
                SingleValueClauseCompare = CompareSingleValuesBoolean;
                IsWithinClauseCompare = CompareIsWithinBoolean;
                SingleValueToIsStmtClauseCompare = CompareSingleValueToIsStmtBoolean;
                IsStmtToIsStmtClauseCompare = CompareIsStmtToIsStmtBoolean;
                RangeToIsStmtClauseCompare = CompareRangeToIsStmtBoolean;
                StandardValueToValueCompare = SimpleCompareBoolean;
            }
            else
            {
                SingleValueClauseCompare = CompareSingleValues;
                IsWithinClauseCompare = CompareIsWithin;
                SingleValueToIsStmtClauseCompare = CompareSingleValueToIsStmt;
                IsStmtToIsStmtClauseCompare = CompareIsStmtToIsStmt;
                RangeToIsStmtClauseCompare = CompareRangeToIsStmt;
                StandardValueToValueCompare = SimpleCompareAny;
            }
        }

        private IdentifierReference GetTheRangeClauseReference(ParserRuleContext rangeClauseCtxt, string theName)
        {
            var simpleNameExpr = ParserRuleContextHelper.GetChild<VBAParser.SimpleNameExprContext>(rangeClauseCtxt);

            var allRefs = new List<IdentifierReference>();
            foreach (var dec in _state.DeclarationFinder.MatchName(theName))
            {
                allRefs.AddRange(dec.References);
            }

            if (!allRefs.Any())
            {
                return null;
            }

            if(allRefs.Count == 1)
            {
                return allRefs.First();
            }
            else
            {
                var rangeClauseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, rangeClauseCtxt)
                                        && (ParserRuleContextHelper.HasParent(rf.Context, simpleNameExpr.Parent)));

                Debug.Assert(rangeClauseReference.Count() == 1);
                return rangeClauseReference.First();
            }
        }

        private string GetRangeClauseText(VBAParser.RangeClauseContext ctxt)
        {
            VBAParser.RelationalOpContext relationalOpCtxt;
            if (TryGetExprContext(ctxt, out relationalOpCtxt))
            {
                return GetTextForRelationalOpContext(relationalOpCtxt);
            }

            var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(ctxt);
            if (smplName != null)
            {
                var rangeClauseIdentifierReference = GetTheRangeClauseReference(smplName, smplName.GetText());
                if (rangeClauseIdentifierReference != null)
                {
                    if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant))
                    {
                        var valuedDeclaration = (ConstantDeclaration)rangeClauseIdentifierReference.Declaration;
                        return valuedDeclaration.Expression;
                    }
                }
            }

            VBAParser.UnaryMinusOpContext negativeCtxt;
            if (TryGetExprContext(ctxt, out negativeCtxt))
            {
                return negativeCtxt.GetText();
            }
            else
            {
                VBAParser.LiteralExprContext theValCtxt;
                if (TryGetExprContext(ctxt, out theValCtxt))
                {
                    var result = GetText(theValCtxt);
                    if (isBooleanType)
                    {
                        if (result.Equals("True"))
                        {
                            return "-1";
                        }
                        else if (result.Equals("False"))
                        {
                            return "0";
                        }
                        else
                        {
                            return result;
                        }
                    }
                    else
                    {
                        return result;
                    }
                }
                else
                {
                    return string.Empty;
                }
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

            if(lExprCtxtIndex != -1 && literalExprCtxtIndex != -1)
            {
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
            }

            var lExprCtxtIndices = new List<int>();
            for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
            {
                var text = relationalOpCtxt.children[idx].GetText();
                if (relationalOpCtxt.children[idx] is VBAParser.LExprContext)
                {
                    lExprCtxt = (VBAParser.LExprContext)relationalOpCtxt.children[idx];
                    lExprCtxtIndices.Add(idx);
                }
                else if (IsComparisonOperator(text))
                {
                    _compareSymbol = text;
                }
            }
            if (lExprCtxtIndices.Count() == 2)
            {
                var ctxt1 = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.First()];
                var expr1 = GetText(ctxt1);

                var ctxt2 = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.Last()];
                var expr2 = GetText(ctxt2);

                if (expr1.Equals(_theRef.IdentifierName))
                {
                    _usesIsClause = true;
                    var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(ctxt2);
                    if (smplName != null)
                    {
                        var rangeClauseIdentifierReference = GetTheRangeClauseReference(smplName, smplName.GetText());
                        if (rangeClauseIdentifierReference != null)
                        {
                            if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant))
                            {
                                var valuedDeclaration = (ConstantDeclaration)rangeClauseIdentifierReference.Declaration;
                                return valuedDeclaration.Expression;
                            }
                        }
                    }
                }
                else if (expr2.Equals(_theRef.IdentifierName))
                {
                    _usesIsClause = true;
                    _compareSymbol = _comparisonOperatorsAndInversions[_compareSymbol];
                    var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(ctxt1);
                    if (smplName != null)
                    {
                        var rangeClauseIdentifierReference = GetTheRangeClauseReference(smplName, smplName.GetText());
                        if (rangeClauseIdentifierReference != null)
                        {
                            if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant))
                            {
                                var valuedDeclaration = (ConstantDeclaration)rangeClauseIdentifierReference.Declaration;
                                return valuedDeclaration.Expression;
                            }
                        }
                    }
                }
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

        private void CheckCaseElseIsReachable(IRangeClause prior)
        {
            if (isBooleanType && !(prior is RangeClauseExtent<int>))
            {
                if (StringToBool(prior.ValueAsString) != StringToBool(ValueAsString))
                {
                    HasUnreachableCaseElse = true;
                }
            }
        }

        //For RangeClauses, a comparison == 0 means that the two range contexts 
        //include all of the same (IsFullyEquivalent) or some of the 
        //same (PartiallyEquivalent) values.

        //The obj passed is always a range clause that is applied prior to the 
        //'this' range clause.  We are checking that 'this' Range Clause is not 
        //made unreachable by the 'prior' RangeClause.
        private int CompareTo(IRangeClause prior)
        {
            if (IsSingleVal && prior.IsSingleVal)
            {
                if (!UsesIsClause && !prior.UsesIsClause)
                {
                    var result = SingleValueClauseCompare(prior.ValueAsString, ValueAsString, EQ);

                    IsFullyEquivalent = result == 0;

                    if (isBooleanType)
                    {
                        CheckCaseElseIsReachable(prior);
                    }

                    return result;
                }
                else if (!UsesIsClause && prior.UsesIsClause)
                {
                    var result = SingleValueClauseCompare(prior.ValueAsString, ValueAsString, prior.CompareSymbol);

                    IsFullyEquivalent = (result == 0);
                    if (isBooleanType)
                    {
                        CheckCaseElseIsReachable(prior);
                    }
                    else if (prior.CompareSymbol.Equals(NEQ))
                    {
                        if (SingleValueClauseCompare(ValueAsString, prior.ValueAsString, EQ) == 0)
                        {
                            HasUnreachableCaseElse = true;
                        }
                    }

                    return result;
                }
                else if (UsesIsClause && !prior.UsesIsClause)
                {
                    var result = SingleValueClauseCompare(ValueAsString, prior.ValueAsString, CompareSymbol);

                    IsPartiallyEquivalent = (result == 0);
                    if (isBooleanType)
                    {
                        CheckCaseElseIsReachable(prior);
                    }
                    else if (CompareSymbol.Equals(NEQ))
                    {
                        if(SingleValueClauseCompare(ValueAsString, prior.ValueAsString, EQ) == 0)
                        {
                            HasUnreachableCaseElse = true;
                        }
                    }

                    return result;
                }
                else if (UsesIsClause && prior.UsesIsClause)
                {
                    var result = IsStmtToIsStmtClauseCompare(prior.ValueAsString, ValueAsString, prior.CompareSymbol, CompareSymbol);

                    if (isBooleanType)
                    {
                        CheckCaseElseIsReachable(prior);
                    }
                    else if (GetOperatorInverseStrict( prior.CompareSymbol).Equals(CompareSymbol)
                        || GetOperatorInverseExtendedSet(prior.CompareSymbol).Equals(CompareSymbol)
                        )
                    {
                        if ((prior.CompareSymbol.Equals(LT) || prior.CompareSymbol.Equals(LTE)) && SingleValueClauseCompare(prior.ValueAsString, ValueAsString, EQ) > 0)
                        {
                            HasUnreachableCaseElse = true;
                        }
                        else if ((prior.CompareSymbol.Equals(GT) || prior.CompareSymbol.Equals(GTE)) && SingleValueClauseCompare(prior.ValueAsString, ValueAsString, EQ) < 0)
                        {
                            HasUnreachableCaseElse = true;
                        }
                        else if ((prior.CompareSymbol.Equals(EQ) || prior.CompareSymbol.Equals(NEQ)) && SingleValueClauseCompare(prior.ValueAsString, ValueAsString, EQ) == 0)
                        {
                            HasUnreachableCaseElse = true;
                        }
                    }

                    return result;
                }
            }
            else if (IsSingleVal && prior.IsRange)
            {
                if (!UsesIsClause)
                {
                    var result = IsWithinClauseCompare(ValueAsString, prior.ValueMinAsString, prior.ValueMaxAsString);
                    IsFullyEquivalent = result == 0;

                    if (isBooleanType)
                    {
                        if (StringToBool(prior.ValueMinAsString)  != StringToBool(prior.ValueMaxAsString))
                        {
                            HasUnreachableCaseElse = true;
                        }
                        else if (StringToBool(ValueAsString) != (StringToBool(prior.ValueMinAsString) || StringToBool(prior.ValueMaxAsString)))
                        {
                            HasUnreachableCaseElse = true;
                        }
                    }

                    return result;
                }
                else
                {
                    // e.g. Case Is > 8 prior Case 3 to 10
                    var resultStartVal = SingleValueToIsStmtClauseCompare(ValueAsString, prior.ValueMinAsString, CompareSymbol);
                    var resultEndVal = SingleValueToIsStmtClauseCompare(ValueAsString, prior.ValueMaxAsString, CompareSymbol);

                    return resultStartVal == 0 || resultEndVal == 0 ? 0 : 1;
                }
            }
            else if (IsRange && prior.IsSingleVal)
            {
                if (!prior.UsesIsClause)
                {
                    var result = IsWithinClauseCompare(prior.ValueAsString, ValueMinAsString, ValueMaxAsString);
                    IsPartiallyEquivalent = result == 0;

                    if (isBooleanType)
                    {
                        if (StringToBool(ValueMinAsString) != (StringToBool(prior.ValueAsString)) && !(prior is RangeClauseExtent<int>))
                        {
                            HasUnreachableCaseElse = true;
                        }
                        else if (StringToBool(ValueMaxAsString) != (StringToBool(prior.ValueAsString)) && !(prior is RangeClauseExtent<int>))
                        {
                            HasUnreachableCaseElse = true;
                        }
                        else if (StringToBool(ValueMinAsString) != (StringToBool(ValueMaxAsString)))
                        {
                            HasUnreachableCaseElse = true;
                        }
                    }
                    return result;
                }
                else
                {
                    var result = RangeToIsStmtClauseCompare(prior.ValueMinAsString, ValueMinAsString, ValueMaxAsString, prior.CompareSymbol);
                    if (StringToBool(ValueMinAsString) != (StringToBool(ValueMaxAsString)))
                    {
                        HasUnreachableCaseElse = true;
                    }
                    return result;
                }
            }
            else if (IsRange && prior.IsRange)
            {
                if (IsWithinClauseCompare(ValueMinAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0
                        && IsWithinClauseCompare(ValueMaxAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0)
                {
                    IsFullyEquivalent = true;
                    return 0;
                }
                else
                {
                    IsPartiallyEquivalent = IsWithinClauseCompare(ValueMinAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0
                        || IsWithinClauseCompare(ValueMaxAsString, prior.ValueMinAsString, prior.ValueMaxAsString) == 0;

                    return IsPartiallyEquivalent ? 0 : ValueMaxAsString.CompareTo(prior.ValueMaxAsString);
                }
            }
            Debug.Assert(true, "Unanticipated code path");
            return 1;
        }

        private void SetIsParseable()
        {
            if (isLongType)
            {
                long longValue;
                if (_isRange)
                {
                    IsParseable = long.TryParse(GetText(_rangeContexts.Key), out longValue)
                            && long.TryParse(GetText(_rangeContexts.Value), out longValue);
                }
                else
                {
                    IsParseable = long.TryParse(GetRangeClauseText(_ctxt), out longValue);
                }
                if (!IsParseable)
                {
                    if (GetText(_ctxt).Contains(".") || (IsStringLiteral && !SelectCaseTypeName.Equals("String")))
                    {
                        _evalResults.MatchesSelectCaseTypeName = false;
                    }
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
                int intVal;
                if (_isRange)
                {
                    IsParseable = int.TryParse(_rangeContexts.Key.GetText(), out intVal)
                            && int.TryParse(_rangeContexts.Value.GetText(), out intVal);
                    if (!IsParseable)
                    {
                        IsParseable = (_rangeContexts.Key.GetText().Equals("True") || _rangeContexts.Key.GetText().Equals("False"))
                            || (_rangeContexts.Value.GetText().Equals("False") || _rangeContexts.Value.GetText().Equals("True"));
                    }
                }
                else
                {
                    IsParseable = int.TryParse(GetRangeClauseText(_ctxt), out intVal);
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
            _evalResults.IsParseable = IsParseable;
        }

        private string GetText(ParserRuleContext ctxt)
        {
            var text = ctxt.GetText();
            //if (!IsStringLiteral)
            //{
            //    IsStringLiteral = text.Contains("\"");
            //}
            return text.Replace("\"", "");
        }

        private int SimpleCompareLong(string value1, string value2)
        {
            return long.Parse(value1).CompareTo(long.Parse(value2));
        }

        private int SimpleCompareDouble(string value1, string value2)
        {
            return double.Parse(value1).CompareTo(double.Parse(value2));
        }

        private int SimpleCompareDecimal(string value1, string value2)
        {
            return decimal.Parse(value1).CompareTo(decimal.Parse(value2));
        }

        private int SimpleCompareAny<T>(T value1, T value2) where T : System.IComparable<T>
        {
            return value1.CompareTo(value2);
        }

        private int SimpleCompareBoolean(string value1, string value2)
        {
            bool val1 = false;
            bool val2 = false;
            int result;
            if (int.TryParse(value1, out result))
            {
                val1 = result != 0;
            }
            else
            {
                val1 = value1.Equals("True");
            }

            if (int.TryParse(value2, out result))
            {
                val2 = result != 0;
            }
            else
            {
                val2 = value2.Equals("True");
            }

            return val1.CompareTo(val2);
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
                //return result < 0 ? 0 : result;
            }
            else if (comparisonSymbol.Equals(LTE))
            {
                return result >= 0 ? 0 : result - 1;
                //return result <= 0 ? 0 : result;
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
                returnVal = IsPartiallyEquivalent ? 0 : returnVal;
            }
            else
            {
                IsFullyEquivalent = !(priorIsStmtCompareSymbol.Contains(NEQ) || priorIsStmtCompareSymbol.Contains(EQ));
                returnVal = IsFullyEquivalent ? 0 : returnVal;
            }
            return returnVal;
        }
#endregion

#region IsWithin
        private int CompareIsWithinLong(string toCheck, string min, string max)
        {
            return CompareIsWithin(long.Parse(toCheck), long.Parse(min), long.Parse(max));
        }

        private int CompareIsWithinDouble(string toCheck, string min, string max)
        {
            return CompareIsWithin(double.Parse(toCheck), double.Parse(min), double.Parse(max));
        }

        private int CompareIsWithinDecimal(string toCheck, string min, string max)
        {
            return CompareIsWithin(decimal.Parse(toCheck), decimal.Parse(min), decimal.Parse(max));
        }

        private int CompareIsWithinBoolean(string toCheck, string min, string max)
        {
            return CompareIsWithin(StringToBool(toCheck), StringToBool(min), StringToBool(max));
        }

        private int CompareIsWithin<T>(T toCheck, T min, T max) where T : System.IComparable<T>
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
