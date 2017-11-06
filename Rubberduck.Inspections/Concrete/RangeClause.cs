using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using static Rubberduck.Inspections.Concrete.UnreachableCaseInspection;

namespace Rubberduck.Inspections.Concrete
{
    public struct RangeClauseEvaluationResults
    {
        public bool IsReachable;
        public bool IsParseable;
        public bool IsStringLiteral;
        public bool IsFullyEquivalent;
        public bool IsPartiallyEquivalent;
        public bool IsIndeterminant;
        public bool CausesUnreachableCaseElse;
        public string NativeTypeName;
        public string TargetTypename;
    }

    internal static class CompareSymbols
    {
        public static readonly string EQ = "=";
        public static readonly string NEQ = "<>";
        public static readonly string LT = "<";
        public static readonly string LTE = "<=";
        public static readonly string GT = ">";
        public static readonly string GTE = ">=";
    }

    public class RangeClause : IRangeClause
    {
        private readonly CaseClauseWrapper _parent;
        private readonly VBAParser.RangeClauseContext _ctxt;
        private bool _usesIsClause;
        private bool _isValueRange;
        private string _compareSymbol;
        private string _valueMinAsString;
        private string _valueMaxAsString;
        private string _valueAsString;
        private RangeClauseEvaluationResults _evalResults;

        internal RangeClause(CaseClauseWrapper caseClause, VBAParser.RangeClauseContext ctxt)
        {
            _ctxt = ctxt;
            _parent = caseClause;
            _compareSymbol = DetermineTheComparisonOperator(ctxt);
            _usesIsClause = HasChildToken(ctxt, Tokens.Is);
            _isValueRange = HasChildToken(ctxt, Tokens.To);

            if (_isValueRange)
            {
                var startValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(_ctxt));
                var endValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(_ctxt));

                var endIsGreaterThanStart = CreateVBAValue(startValueAsString) > CreateVBAValue(endValueAsString);
                _valueMinAsString = endIsGreaterThanStart ? startValueAsString : endValueAsString;
                _valueMaxAsString = endIsGreaterThanStart ? endValueAsString : startValueAsString;
                _valueAsString = _valueMinAsString;
            }
            else
            {
                _valueAsString = GetRangeClauseText(_ctxt);
                _valueMinAsString = _valueAsString;
                _valueMaxAsString = _valueAsString;
            }

            _evalResults = new RangeClauseEvaluationResults
            {
                IsFullyEquivalent = false,
                IsPartiallyEquivalent = false,
                IsReachable = true,
                IsStringLiteral = IsStringLiteral,
                IsIndeterminant = false,
                TargetTypename = caseClause.Parent.TypeName,
                IsParseable = CanParseTo(caseClause.Parent.TypeName),
                NativeTypeName = EvaluateRangeClauseTypeName(caseClause)
            };
        }

        public bool IsParseable => _evalResults.IsParseable;
        public bool MatchesSelectCaseType => _evalResults.NativeTypeName.Equals(_evalResults.TargetTypename);
        public string RangeClauseTypeName => _evalResults.NativeTypeName;

        //IRangeClause implementation
        public bool IsSingleVal => !IsRange;
        public bool IsRange => _isValueRange;
        public bool UsesIsClause => _usesIsClause;
        public bool IsRangeExtent => false;
        public string ValueAsString => _valueAsString;
        public string ValueMinAsString => _valueMinAsString;
        public string ValueMaxAsString => _valueMaxAsString;
        public string CompareSymbol => _compareSymbol;

        private VBAValue CreateVBAValue(string AsString) => new VBAValue(AsString, SelectCaseTypeName);
        private string SelectCaseTypeName => _parent.Parent.TypeName;
        private bool IsSelectCaseBoolean => _parent.Parent.TypeName.Equals(Tokens.Boolean);
        private bool IsStringLiteral => _ctxt.GetText().StartsWith("\"") && _ctxt.GetText().EndsWith("\"");

        private string EvaluateRangeClauseTypeName(CaseClauseWrapper caseClause)
        {
            var textValue = _ctxt.GetText();
            if (IsStringLiteral)
            {
                return Tokens.String;
            }
            else if (textValue.EndsWith("#"))
            {
                return Tokens.Double;
            }
            else if (textValue.Contains("."))
            {
                return caseClause.Parent.TypeName.Equals(Tokens.Currency) ? Tokens.Currency : Tokens.Double;
            }
            else if (textValue.Equals(Tokens.True) || textValue.Equals(Tokens.False))
            {
                return Tokens.Boolean;
            }
            else
            {
                return caseClause.Parent.TypeName;
            }
        }

        private IdentifierReference GetTheRangeClauseReference(ParserRuleContext rangeClauseCtxt, string theName)
        {
            var allRefs = new List<IdentifierReference>();
            foreach (var dec in _parent.Parent.State.DeclarationFinder.MatchName(theName))
            {
                allRefs.AddRange(dec.References);
            }

            if (!allRefs.Any())
            {
                return null;
            }

            if (allRefs.Count == 1)
            {
                return allRefs.First();
            }
            else
            {
                var simpleNameExpr = ParserRuleContextHelper.GetChild<VBAParser.SimpleNameExprContext>(rangeClauseCtxt);
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

            VBAParser.LExprContext lExprContext;
            if (TryGetExprContext(ctxt, out lExprContext))
            {
                string expressionValue;
                return TryGetTheExpressionValue(lExprContext, out expressionValue) ? expressionValue : string.Empty;
            }

            VBAParser.UnaryMinusOpContext negativeCtxt;
            if (TryGetExprContext(ctxt, out negativeCtxt))
            {
                return negativeCtxt.GetText();
            }

            VBAParser.LiteralExprContext theValCtxt;
            return TryGetExprContext(ctxt, out theValCtxt) ? GetText(theValCtxt) : string.Empty;
        }

        private string GetTextForRelationalOpContext(VBAParser.RelationalOpContext relationalOpCtxt)
        {
            var lExprCtxtIndices = new List<int>();
            var literalExprCtxtIndices = new List<int>();

            for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
            {
                var text = relationalOpCtxt.children[idx].GetText();
                if (relationalOpCtxt.children[idx] is VBAParser.LExprContext)
                {
                    lExprCtxtIndices.Add(idx);
                }
                else if (relationalOpCtxt.children[idx] is VBAParser.LiteralExprContext)
                {
                    literalExprCtxtIndices.Add(idx);
                }
                else if (RangeClauseComparer.IsComparisonOperator(text))
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

                _usesIsClause = true;
                if (expr1.Equals(_parent.Parent.IdReference.IdentifierName))
                {
                    string result;
                    if (TryGetTheExpressionValue(ctxt2, out result))
                    {
                        return result;
                    }
                }
                else if (expr2.Equals(_parent.Parent.IdReference.IdentifierName))
                {
                    string result;
                    _compareSymbol = RangeClauseComparer.GetCompareSymbolInverse(_compareSymbol);
                    if (TryGetTheExpressionValue(ctxt1, out result))
                    {
                        return result;
                    }
                }
            }
            else if (lExprCtxtIndices.Count == 1 && literalExprCtxtIndices.Count == 1)
            {
                if (lExprCtxtIndices.First() > literalExprCtxtIndices.First())
                {
                    //A greater lExprCtxtIndex means the comparison is of the form '10 < z'...invert 
                    //the operator to make the expression conform to 'z > 10' or 'Is > 10'
                    _compareSymbol = RangeClauseComparer.GetCompareSymbolInverse(_compareSymbol);
                }
                var lExprCtxt = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.First()];
                if (lExprCtxt.GetText().Equals(_parent.Parent.IdReference.IdentifierName))
                {
                    _usesIsClause = true;
                    var theValueCtxt = (VBAParser.LiteralExprContext)relationalOpCtxt.children[literalExprCtxtIndices.First()];
                    return theValueCtxt.GetText();
                }
            }
            return string.Empty;
        }

        private bool TryGetTheExpressionValue(VBAParser.LExprContext ctxt, out string expressionValue)
        {
            expressionValue = string.Empty;
            var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(ctxt);
            if (smplName != null)
            {
                var rangeClauseIdentifierReference = GetTheRangeClauseReference(smplName, smplName.GetText());
                if (rangeClauseIdentifierReference != null)
                {
                    if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant))
                    {
                        var valuedDeclaration = (ConstantDeclaration)rangeClauseIdentifierReference.Declaration;
                        expressionValue = valuedDeclaration.Expression;
                        return true;
                    }
                }
            }
            return false;
        }

        private bool CanParseTo(string targetTypeName)
        {
            var textValues = new List<string>();
            if (_isValueRange)
            {
                var start = ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(_ctxt);
                var end = ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(_ctxt);

                var startVal = CreateVBAValue(GetText(start));
                var endVal = CreateVBAValue(GetText(end));

                return startVal.IsParseableToTypeName(targetTypeName)
                    && endVal.IsParseableToTypeName(targetTypeName);
            }
            else
            {
                return CreateVBAValue(GetRangeClauseText(_ctxt)).IsParseableToTypeName(targetTypeName);
            }
        }

        private string GetText(ParserRuleContext ctxt)
        {
            var text = ctxt.GetText();
            return text.Replace("\"", "");
        }

        private string DetermineTheComparisonOperator(VBAParser.RangeClauseContext ctxt)
        {
            _usesIsClause = false;
            //'VBAParser.ComparisonOperatorContext' - The 'Is' case
            var opCtxt = ParserRuleContextHelper.GetChild<VBAParser.ComparisonOperatorContext>(ctxt);
            if (opCtxt != null)
            {
                _usesIsClause = true;
                return opCtxt.GetText();
            }
            return CompareSymbols.EQ;
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

    public class RangeClauseComparer
    {

        public struct RangeCompareData
        {
            public IRangeClause Current;
            public IRangeClause Prior;
            public string CurrentCompareSymbol;
            public string PriorCompareSymbol;
            public VBAValue CurrentValue;
            public VBAValue PriorValue;
            public VBAValue CurrentValueMin;
            public VBAValue CurrentValueMax;
            public VBAValue PriorValueMin;
            public VBAValue PriorValueMax;
            public string SelectCaseTypename;

            public RangeCompareData(IRangeClause current, IRangeClause prior, string targetTypeName)
            {
                Current = current;
                Prior = prior;
                CurrentCompareSymbol = current.CompareSymbol;
                PriorCompareSymbol = prior.CompareSymbol;
                SelectCaseTypename = targetTypeName;

                CurrentValue = new VBAValue(current.ValueAsString, SelectCaseTypename);
                CurrentValueMin = current.IsSingleVal ? CurrentValue : new VBAValue(current.ValueMinAsString, SelectCaseTypename);
                CurrentValueMax = current.IsSingleVal ? CurrentValue : new VBAValue(current.ValueMaxAsString, SelectCaseTypename);

                PriorValue = new VBAValue(prior.ValueAsString, SelectCaseTypename);
                PriorValueMin = prior.IsSingleVal ? PriorValue : new VBAValue(prior.ValueMinAsString, SelectCaseTypename);
                PriorValueMax = prior.IsSingleVal ? PriorValue : new VBAValue(prior.ValueMaxAsString, SelectCaseTypename);
            }
        }

        private static Dictionary<string, string> _compareInversions = new Dictionary<string, string>()
        {
            { CompareSymbols.EQ, CompareSymbols.NEQ },
            { CompareSymbols.NEQ, CompareSymbols.EQ },
            { CompareSymbols.LT, CompareSymbols.GT },
            { CompareSymbols.LTE, CompareSymbols.GTE },
            { CompareSymbols.GT, CompareSymbols.LT },
            { CompareSymbols.GTE, CompareSymbols.LTE }
        };

        private static Dictionary<string, string> _compareInversionsExtended = new Dictionary<string, string>()
        {
            { CompareSymbols.EQ, CompareSymbols.NEQ },
            { CompareSymbols.NEQ, CompareSymbols.EQ },
            { CompareSymbols.LT, CompareSymbols.GTE },
            { CompareSymbols.LTE, CompareSymbols.GT },
            { CompareSymbols.GT, CompareSymbols.LTE },
            { CompareSymbols.GTE, CompareSymbols.LT }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> CompareOps = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { CompareSymbols.EQ, delegate(VBAValue LHS, VBAValue RHS){ return LHS == RHS; } },
            { CompareSymbols.NEQ, delegate(VBAValue LHS, VBAValue RHS){ return LHS != RHS; } },
            { CompareSymbols.LT, delegate(VBAValue LHS, VBAValue RHS){ return LHS < RHS; } },
            { CompareSymbols.LTE, delegate(VBAValue LHS, VBAValue RHS){ return LHS <= RHS; } },
            { CompareSymbols.GT, delegate(VBAValue LHS, VBAValue RHS){ return LHS > RHS; } },
            { CompareSymbols.GTE, delegate(VBAValue LHS, VBAValue RHS){ return LHS >= RHS; } }
        };

        private string _targetTypeName;
        private bool IsBooleanSelectCase => _targetTypeName.Equals(Tokens.Boolean);

        public bool IsFullyEquivalent { set; get; }
        public bool IsPartiallyEquivalent { set; get; }
        public bool CausesUnreachableCaseElse { set; get; }
        public bool IsReachable => !IsFullyEquivalent && !IsPartiallyEquivalent;

        public static bool IsComparisonOperator(string opCandidate) => _compareInversions.Keys.Contains(opCandidate);
        public static string GetCompareSymbolInverse(string theOperator)
        {
            return IsComparisonOperator(theOperator) ? _compareInversions[theOperator] : theOperator;
        }

        public void Compare(IRangeClause current, IRangeClause prior, string targetTypeName)
        {
            _targetTypeName = targetTypeName;
            IsFullyEquivalent = false;
            IsPartiallyEquivalent = false;

            var dto = new RangeCompareData(current, prior, targetTypeName);

            if (prior.IsRangeExtent)
            {
                if (current.IsSingleVal)
                {
                    IsFullyEquivalent = CompareOps[dto.PriorCompareSymbol](dto.CurrentValue, dto.PriorValue);
                }
            }
            else if (current.IsSingleVal && prior.IsSingleVal)
            {
                CompareSingleValues(dto);
            }
            else if (current.IsRange || prior.IsRange)
            {
                CompareRangeValues(dto);
            }
        }

        private void CompareSingleValues(RangeCompareData dto)
        {
            if (!dto.Current.UsesIsClause && !dto.Prior.UsesIsClause)
            {
                IsFullyEquivalent = dto.CurrentValue == dto.PriorValue;

                CausesUnreachableCaseElse = IsBooleanSelectCase && !IsFullyEquivalent;
            }
            else if (!dto.Current.UsesIsClause && dto.Prior.UsesIsClause)
            {
                var compareOP = CompareOps[dto.PriorCompareSymbol];
                IsFullyEquivalent = compareOP(dto.CurrentValue, dto.PriorValue);

                if (IsBooleanSelectCase)
                {
                    CausesUnreachableCaseElse = !IsFullyEquivalent;
                }
                else if (dto.PriorCompareSymbol.Equals(CompareSymbols.NEQ))
                {
                    CausesUnreachableCaseElse = dto.CurrentValue == dto.PriorValue;
                }
            }
            else if (dto.Current.UsesIsClause && !dto.Prior.UsesIsClause)
            {
                var compareOP = CompareOps[GetCompareSymbolInverse(dto.CurrentCompareSymbol)];
                IsPartiallyEquivalent = compareOP(dto.CurrentValue, dto.PriorValue);
                if (IsBooleanSelectCase)
                {
                    CausesUnreachableCaseElse = dto.CurrentValue != dto.PriorValue;
                }
                else if (dto.CurrentCompareSymbol.Equals(GetCompareSymbolInverse(dto.PriorCompareSymbol)))
                {
                    CausesUnreachableCaseElse = compareOP(dto.CurrentValue, dto.PriorValue);
                }
            }
            else if (dto.Current.UsesIsClause && dto.Prior.UsesIsClause)
            {
                CompareIsStmtToPriorIsStmt(dto);
            }
        }

        private void CompareRangeValues(RangeCompareData dto)
        {
            if (dto.Current.IsSingleVal && dto.Prior.IsRange)
            {
                CompareSingleValueToPriorRange(dto);
                if (IsBooleanSelectCase)
                {
                    if (dto.PriorValueMin.AsBoolean() != dto.PriorValueMax.AsBoolean())
                    {
                        CausesUnreachableCaseElse = true;
                    }
                    else if (dto.CurrentValue.AsBoolean() != (dto.PriorValueMin.AsBoolean().Value || dto.PriorValueMax.AsBoolean().Value))
                    {
                        CausesUnreachableCaseElse = true;
                    }
                }
            }
            if (dto.Current.IsRange && dto.Prior.IsSingleVal)
            {
                CompareRangeToPriorSingleValue(dto);
            }
            if (dto.Current.IsRange && dto.Prior.IsRange)
            {
                CompareRangeToPriorRange(dto);
            }
        }

        private void CompareSingleValueToPriorRange(RangeCompareData dto)
        {
            if (dto.Current.UsesIsClause)
            {
                // e.g. Case Is > 8 prior Case 3 to 10
                if (dto.CurrentValue.IsWithin(dto.PriorValueMin, dto.PriorValueMax))
                {
                    IsPartiallyEquivalent = true;
                    return;
                }

                // e.g. Current Case: Is > 2 Prior Case: 3 to 10
                if (dto.CurrentValue < dto.PriorValueMin)
                {
                    if (dto.CurrentCompareSymbol.Equals(CompareSymbols.GT) || dto.CurrentCompareSymbol.Equals(CompareSymbols.GTE))
                    {
                        IsPartiallyEquivalent = true;
                        return;
                    }
                }
                // e.g. Current Case Is < 15 Prior Case 3 to 10
                if (dto.CurrentValue > dto.PriorValueMax)
                {
                    if (dto.CurrentCompareSymbol.Equals(CompareSymbols.LT) || dto.CurrentCompareSymbol.Equals(CompareSymbols.LTE))
                    {
                        IsPartiallyEquivalent = true;
                        return;
                    }
                }
            }
            else
            {
                IsFullyEquivalent = dto.CurrentValue.IsWithin(dto.PriorValueMin, dto.PriorValueMax);
            }
        }

        private void CompareRangeToPriorSingleValue(RangeCompareData dto)
        {
            if (!dto.Prior.UsesIsClause)
            {
                IsPartiallyEquivalent = dto.PriorValue.IsWithin(dto.CurrentValueMin, dto.CurrentValueMax);
                IsFullyEquivalent = false;

                if (IsBooleanSelectCase)
                {
                    CausesUnreachableCaseElse = dto.CurrentValue != dto.PriorValue;
                }
            }
            else  //prior uses Is Clause
            {
                CompareRangeToIsStmtExt(dto.PriorValue, dto.CurrentValueMin, dto.CurrentValueMax, dto.PriorCompareSymbol);
                if (IsBooleanSelectCase)
                {
                    CausesUnreachableCaseElse = dto.CurrentValueMin != dto.CurrentValueMax;
                }
            }
        }

        private void CompareRangeToPriorRange(RangeCompareData dto)
        {
            IsFullyEquivalent = dto.CurrentValueMin.IsWithin(dto.PriorValueMin, dto.PriorValueMax)
                    && dto.CurrentValueMax.IsWithin(dto.PriorValueMin, dto.PriorValueMax);

            if(!IsFullyEquivalent)
            {
                IsPartiallyEquivalent = dto.CurrentValueMin.IsWithin(dto.PriorValueMin, dto.PriorValueMax)
                    || dto.CurrentValueMax.IsWithin(dto.PriorValueMin, dto.PriorValueMax);
            }
        }

        private void CompareSingleValueToPriorIsClause(RangeCompareData dto)
        {
            var compareOP = CompareOps[dto.PriorCompareSymbol];
            IsFullyEquivalent = compareOP(dto.CurrentValue, dto.PriorValue);
        }

        private void CompareIsClauseToPriorSingleValue(RangeCompareData dto)
        {
            //e.g. Current Case Is > 9, Prior Case 10
            var compareOP = CompareOps[GetCompareSymbolInverse( dto.CurrentCompareSymbol )];
            IsPartiallyEquivalent = compareOP(dto.CurrentValue, dto.PriorValue);
        }

        private void CompareIsStmtToPriorIsStmt(RangeCompareData dto)
        {
            /*
             * Current Case Is < 5 Prior Is < 10 : Fully
             * Current Case Is > 5 Prior Is > 10 : Partial
             * Current Case Is < 10 Prior Is < 5 : Partial
             * Current Case Is > 10 Prior Is > 5 : Fully
             */
            if (dto.CurrentCompareSymbol.Equals(dto.PriorCompareSymbol))
            {
                var compareOp = CompareOps[dto.CurrentCompareSymbol];
                IsFullyEquivalent = compareOp(dto.CurrentValue, dto.PriorValue);
                if (!IsFullyEquivalent)
                {
                    compareOp = CompareOps[GetCompareSymbolInverse(dto.CurrentCompareSymbol)];
                    IsPartiallyEquivalent = compareOp(dto.CurrentValue, dto.PriorValue);
                }
            }
            /*
            * Current Case Is > 5 Prior Is < 10 : Partial
            * Current Case Is < 5 Prior Is > 10 : No conflict
            * Current Case Is > 10 Prior Is < 5 : No conflict
            * Current Case Is < 10 Prior Is > 5 : Partial
            * */
            else
            {
                var compareOp = CompareOps[GetCompareSymbolInverse(dto.CurrentCompareSymbol)];
                IsPartiallyEquivalent = compareOp(dto.CurrentValue, dto.PriorValue);
            }

            if(dto.CurrentCompareSymbol.Equals(GetCompareSymbolInverse(dto.PriorCompareSymbol))
                || dto.CurrentCompareSymbol.Equals(_compareInversionsExtended[dto.PriorCompareSymbol]))
            {
                CausesUnreachableCaseElse = IsPartiallyEquivalent 
                    || dto.CurrentCompareSymbol.Equals(CompareSymbols.EQ) || dto.PriorCompareSymbol.Equals(CompareSymbols.EQ);
            }
        }

        private void CompareRangeToIsStmtExt(VBAValue priorIsStmtValue, VBAValue minVal, VBAValue maxVal, string priorCompareSymbol)
        {
            if (priorCompareSymbol.Equals(CompareSymbols.GT) || priorCompareSymbol.Equals(CompareSymbols.GTE))
            {
                IsFullyEquivalent = priorCompareSymbol.Equals(CompareSymbols.GT) ? minVal > priorIsStmtValue : minVal >= priorIsStmtValue;
                if (!IsFullyEquivalent)
                {
                    IsPartiallyEquivalent = priorCompareSymbol.Equals(CompareSymbols.GT) ? maxVal > priorIsStmtValue : maxVal >= priorIsStmtValue;
                }
            }
            else if (priorCompareSymbol.Equals(CompareSymbols.LT) || priorCompareSymbol.Equals(CompareSymbols.LTE))
            {
                IsFullyEquivalent = priorCompareSymbol.Equals(CompareSymbols.LT) ? maxVal < priorIsStmtValue : maxVal <= priorIsStmtValue;
                if (!IsFullyEquivalent)
                {
                    IsPartiallyEquivalent = priorCompareSymbol.Equals(CompareSymbols.LT) ? minVal < priorIsStmtValue : minVal <= priorIsStmtValue;
                }
            }
        }
    }
    #region VBAValue
    public class VBAValue
    {
        private readonly string _targetTypeName;
        private readonly string _valueAsString;
        private readonly Func<VBAValue, VBAValue, bool> _operatorGT;
        private readonly Func<VBAValue, VBAValue, bool> _operatorLT;
        private readonly Func<VBAValue, VBAValue, bool> _operatorEQ;

        private long? _valueAsLong;
        private double? _valueAsDouble;
        private decimal? _valueAsDecimal;
        private bool? _valueAsBoolean;

        private long resultLong;
        private double resultDouble;
        private decimal resultDecimal;


        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> OperatorsGT = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Long, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Double, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsCurrency().Value > compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : !thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) > 0; } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> OperatorsLT = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Long, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Double, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsCurrency().Value < compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) < 0; } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, bool>> OperatorsEQ = new Dictionary<string, Func<VBAValue, VBAValue, bool>>()
        {
            { Tokens.Integer, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Long, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Double, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsCurrency().Value == compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value; } },
            { Tokens.String, delegate(VBAValue thisValue, VBAValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) == 0; } }
        };

        private static Dictionary<string, Func<VBAValue, bool>> IsParseableTests = new Dictionary<string, Func<VBAValue, bool>>()
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

        public VBAValue(string valueToken, string targetTypeName)
        {
            _valueAsString = valueToken.EndsWith("#") ? valueToken.Replace("#", ".00") : valueToken;
            _targetTypeName = targetTypeName;

            Debug.Assert(OperatorsGT.ContainsKey(targetTypeName));
            Debug.Assert(OperatorsLT.ContainsKey(targetTypeName));
            Debug.Assert(OperatorsEQ.ContainsKey(targetTypeName));

            _operatorGT = OperatorsGT[targetTypeName];
            _operatorLT = OperatorsLT[targetTypeName];
            _operatorEQ = OperatorsEQ[targetTypeName];
        }

        public string TargetTypeName => _targetTypeName;

        public bool IsParseableToTypeName(string typeName) => IsParseableTests.ContainsKey(typeName) ? IsParseableTests[typeName](this) : false;

        public bool IsWithin(VBAValue start, VBAValue end ) 
            => start > end ? this >= end && this <= start : this >= start && this <= end;

        public static bool operator >(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._operatorGT(thisValue, compValue);
        }

        public static bool operator <(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._operatorLT(thisValue, compValue);
        }

        public static bool operator ==(VBAValue thisValue, VBAValue compValue)
        {
            return thisValue._operatorEQ(thisValue, compValue);
        }

        public static bool operator !=(VBAValue thisValue, VBAValue compValue)
        {
            return !thisValue._operatorEQ(thisValue, compValue);
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
            if (obj == null || !(obj is VBAValue))
            {
                return false;
            }
            var asValue = (VBAValue)obj;
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
#endregion
}
