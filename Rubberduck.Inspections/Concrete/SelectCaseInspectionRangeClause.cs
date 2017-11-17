using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using static Rubberduck.Inspections.Concrete.SelectCaseInspection;

namespace Rubberduck.Inspections.Concrete
{
    public struct CompareResults
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

    public class SelectCaseInspectionRangeClause
    {
        private readonly CaseClauseWrapper _parent;
        private readonly VBAParser.RangeClauseContext _ctxt;
        private bool _usesIsClause;
        private bool _isValueRange;
        private bool _isRangeExtent;
        private string _compareSymbol;
        private string _valueMinAsString;
        private string _valueMaxAsString;
        private string _valueAsString;
        private CompareResults _evalResults;

        internal SelectCaseInspectionRangeClause(CaseClauseWrapper caseClause, VBAParser.RangeClauseContext ctxt)
        {
            _ctxt = ctxt;
            _parent = caseClause;
            _usesIsClause = HasChildToken(ctxt, Tokens.Is);
            _isValueRange = HasChildToken(ctxt, Tokens.To);
            _compareSymbol = _usesIsClause ? GetTheCompareOperator(ctxt) : CompareSymbols.EQ;

            _isRangeExtent = false;

            if (_isValueRange)
            {
                var startValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(_ctxt));
                var endValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(_ctxt));

                var endIsGreaterThanStart = CreateVBAValue(endValueAsString) > CreateVBAValue(startValueAsString);
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

            _evalResults = new CompareResults
            {
                IsFullyEquivalent = false,
                IsPartiallyEquivalent = false,
                IsReachable = true,
                IsStringLiteral = IsStringLiteral(ctxt.GetText()),
                IsIndeterminant = false,
                TargetTypename = caseClause.Parent.TypeName,
                IsParseable = CanParseTo(caseClause.Parent.TypeName),
                NativeTypeName = EvaluateRangeClauseTypeName(caseClause)
            };

            _evalResults.IsIndeterminant = !IsParseable && _evalResults.NativeTypeName.Equals(_evalResults.TargetTypename);
        }

        public static SelectCaseInspectionRangeClause CreateBoundaryCheckRangeClause(string boundaryValue, string compareSymbol)
        {
            return new SelectCaseInspectionRangeClause(boundaryValue, compareSymbol);
        }

        private SelectCaseInspectionRangeClause(string typeBoundary, string compareSymbol)
        {
            _isRangeExtent = true;
            _isValueRange = false;
            _usesIsClause = true;
            _valueAsString = typeBoundary;
            _valueMinAsString = typeBoundary;
            _valueMinAsString = typeBoundary;
            _compareSymbol = compareSymbol;
        }

        public bool IsParseable => _evalResults.IsParseable;
        public bool CompareByTextOnly => _evalResults.IsIndeterminant;
        public bool MatchesSelectCaseType => _evalResults.NativeTypeName.Equals(_evalResults.TargetTypename);
        public string RangeClauseTypeName => _evalResults.NativeTypeName;
        public VBAParser.RangeClauseContext Context => _ctxt;

        public bool IsSingleVal => !IsRange;
        public bool IsRange => _isValueRange;
        public bool UsesIsClause => _usesIsClause;
        public bool IsRangeExtent => _isRangeExtent;
        public string ValueAsString => _valueAsString;
        public string ValueMinAsString => _valueMinAsString;
        public string ValueMaxAsString => _valueMaxAsString;
        public string CompareSymbol => _compareSymbol;

        private VBAValue CreateVBAValue(string AsString) => new VBAValue(AsString, SelectCaseTypeName);
        private string SelectCaseTypeName => _parent.Parent.TypeName;
        private bool IsSelectCaseBoolean => _parent.Parent.TypeName.Equals(Tokens.Boolean);
        private bool IsStringLiteral(string text) => text.StartsWith("\"") && _ctxt.GetText().EndsWith("\"");

        private string EvaluateRangeClauseTypeName(CaseClauseWrapper caseClause)
        {
            var textValue = _ctxt.GetText();
            if (IsStringLiteral(textValue))
            {
                return Tokens.String;
            }
            else if (textValue.EndsWith("#"))
            {
                var modified = textValue.Substring(0,textValue.Length -1);
                long LHS;
                if (long.TryParse(modified,out LHS))
                {
                    return Tokens.Double;
                }
                return Tokens.String;
            }
            else if (textValue.Contains("."))
            {
                double result;
                if(double.TryParse(textValue, out result))
                {
                    return Tokens.Double;
                }

                decimal currency;
                if (decimal.TryParse(textValue, out currency))
                {
                    return Tokens.Currency;
                }
                return caseClause.Parent.TypeName;
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
            if (TryGetChildContext(ctxt, out relationalOpCtxt))
            {
                _usesIsClause = true;
                return GetTextForRelationalOpContext(relationalOpCtxt);
            }

            VBAParser.LExprContext lExprContext;
            if (TryGetChildContext(ctxt, out lExprContext))
            {
                string expressionValue;
                return TryGetTheExpressionValue(lExprContext, out expressionValue) ? expressionValue : string.Empty;
            }

            VBAParser.UnaryMinusOpContext negativeCtxt;
            if (TryGetChildContext(ctxt, out negativeCtxt))
            {
                return negativeCtxt.GetText();
            }

            VBAParser.LiteralExprContext theValCtxt;
            return TryGetChildContext(ctxt, out theValCtxt) ? GetText(theValCtxt) : string.Empty;
        }

        private string GetTextForRelationalOpContext(VBAParser.RelationalOpContext relationalOpCtxt)
        {
            var lExprCtxtIndices = new List<int>();
            var literalExprCtxtIndices = new List<int>();
           // _usesIsClause = true;

            for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
            {
                var text = relationalOpCtxt.children[idx].GetText();
                if (relationalOpCtxt.children[idx] is VBAParser.LExprContext)
                {
                    lExprCtxtIndices.Add(idx);
                }
                else if (relationalOpCtxt.children[idx] is VBAParser.UnaryMinusOpContext
                        || relationalOpCtxt.children[idx] is VBAParser.LiteralExprContext)
                {
                    literalExprCtxtIndices.Add(idx);
                }
                else if (RangeClauseComparer.IsComparisonOperator(text))
                {
                    _compareSymbol = text;
                }
            }

            if (lExprCtxtIndices.Count() == 2)  //e.g., x > someConstantExpression
            {
                var ctxtLHS = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.First()];
                var ctxtRHS = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.Last()];

                string result;
                if (GetText(ctxtLHS).Equals(_parent.Parent.IdReference.IdentifierName))
                {
                    return TryGetTheExpressionValue(ctxtRHS, out result) ? result : string.Empty;
                }
                else if (GetText(ctxtRHS).Equals(_parent.Parent.IdReference.IdentifierName))
                {
                    _compareSymbol = RangeClauseComparer.GetInverse(_compareSymbol);
                    return TryGetTheExpressionValue(ctxtLHS, out result) ? result : string.Empty;
                }
            }
            else if (lExprCtxtIndices.Count == 1 && literalExprCtxtIndices.Count == 1) // e.g., z < 10
            {
                var lExpIndex = lExprCtxtIndices.First();
                var litExpIndex = literalExprCtxtIndices.First();
                var lExprCtxt = (VBAParser.LExprContext)relationalOpCtxt.children[lExpIndex];
                if (GetText(lExprCtxt).Equals(_parent.Parent.IdReference.IdentifierName))
                {
                    _compareSymbol = lExpIndex > litExpIndex ? 
                        RangeClauseComparer.GetInverse(_compareSymbol) : _compareSymbol;
                    return GetText((ParserRuleContext)relationalOpCtxt.children[litExpIndex]);
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

        private string GetTheCompareOperator(VBAParser.RangeClauseContext ctxt)
        {
            VBAParser.ComparisonOperatorContext opCtxt;
            _usesIsClause = TryGetChildContext(ctxt, out opCtxt);
            return opCtxt != null ? opCtxt.GetText() : CompareSymbols.EQ;
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

        private static bool TryGetChildContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : ParserRuleContext //VBAParser.ExpressionContext
        {
            opCtxt = null;
            opCtxt = ParserRuleContextHelper.GetChild<U>(ctxt);
            return opCtxt != null;
        }
    }

    public class RangeClauseComparer
    {
        internal struct RangeCompareData
        {
            public SelectCaseInspectionRangeClause Current;
            public SelectCaseInspectionRangeClause Prior;
            public string CurrentCompareSymbol;
            public string PriorCompareSymbol;
            public VBAValue CurrentValue;
            public VBAValue PriorValue;
            public VBAValue CurrentValueMin;
            public VBAValue CurrentValueMax;
            public VBAValue PriorValueMin;
            public VBAValue PriorValueMax;
            public string SelectCaseTypename;

            public RangeCompareData(SelectCaseInspectionRangeClause current, SelectCaseInspectionRangeClause prior, string targetTypeName)
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

        public struct CompareResultData
        {
            public bool IsRedundant;
            public bool HasConflict;
            public bool MakesAllRemainingCasesUnreachable;
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
        private CompareResultData _results;

        public bool IsFullyEquivalent { set; get; }
        public bool IsPartiallyEquivalent { set; get; }
        public bool CausesUnreachableCaseElse { set; get; }
        public bool IsReachable => !IsFullyEquivalent && !IsPartiallyEquivalent;

        public static bool IsComparisonOperator(string opCandidate) => _compareInversions.Keys.Contains(opCandidate);
        public static string GetInverse(string theOperator)
        {
            return IsComparisonOperator(theOperator) ? _compareInversions[theOperator] : theOperator;
        }

        public CompareResultData Compare(SelectCaseInspectionRangeClause current, SelectCaseInspectionRangeClause prior, string targetTypeName)
        {
            _targetTypeName = targetTypeName;
            IsFullyEquivalent = false;
            IsPartiallyEquivalent = false;

            var dto = new RangeCompareData(current, prior, targetTypeName);

            if (prior.IsRangeExtent)
            {
                if (current.IsSingleVal)
                {
#if (DEBUG)
                    var currentVal = dto.CurrentValue.AsString();
                    var priorVal = dto.PriorValue.AsString();
                    var priorSymbol = dto.PriorCompareSymbol;
#endif
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
            _results.IsRedundant = IsFullyEquivalent;
            _results.HasConflict = IsPartiallyEquivalent;
            _results.MakesAllRemainingCasesUnreachable = CausesUnreachableCaseElse;
            return _results;
        }

        private void CompareSingleValuesSimple(RangeCompareData dto)
        {
            IsFullyEquivalent = dto.CurrentValue == dto.PriorValue;

            CausesUnreachableCaseElse = IsBooleanSelectCase && !IsFullyEquivalent;
        }

        private void CompareSingleValues(RangeCompareData dto)
        {
            var current = dto.CurrentValue.AsLong();
            var prior = dto.PriorValue.AsLong();

            if (!dto.Current.UsesIsClause && !dto.Prior.UsesIsClause)
            {
                CompareSingleValuesSimple(dto);
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

                if (dto.CurrentCompareSymbol.Equals(CompareSymbols.EQ))
                {
                    CompareSingleValuesSimple(dto);
                }
                else if (dto.CurrentCompareSymbol.Equals(CompareSymbols.NEQ))
                {
                    IsPartiallyEquivalent = dto.CurrentValue != dto.PriorValue;

                    CausesUnreachableCaseElse = !IsPartiallyEquivalent;
                }
                else
                {
                    var compareOP = CompareOps[GetInverse(dto.CurrentCompareSymbol)];
                    IsPartiallyEquivalent = compareOP(dto.CurrentValue, dto.PriorValue);
                }
                if (IsBooleanSelectCase)
                {
                    CausesUnreachableCaseElse = dto.CurrentValue != dto.PriorValue;
                }
                else if (dto.CurrentCompareSymbol.Equals(GetInverse(dto.PriorCompareSymbol)))
                {
                    var compareOP = CompareOps[GetInverse(dto.CurrentCompareSymbol)];
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
            var compareOP = CompareOps[GetInverse( dto.CurrentCompareSymbol )];
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
                    compareOp = CompareOps[GetInverse(dto.CurrentCompareSymbol)];
                    IsPartiallyEquivalent = compareOp(dto.CurrentValue, dto.PriorValue);
                }
            }
            /*
            * Current Case Is = 5 Prior Is < 10 : Fully
            * Current Case Is > 5 Prior Is < 10 : Partial
            * Current Case Is < 5 Prior Is > 10 : No conflict
            * Current Case Is > 10 Prior Is < 5 : No conflict
            * Current Case Is < 10 Prior Is > 5 : Partial
            * Current Case Is < 10 Prior Is = 5 : Partial
            * */
            else
            {
                if (dto.CurrentCompareSymbol.Equals(CompareSymbols.EQ))
                {
                    CompareSingleValueToPriorIsClause(dto);
                }
                else if(dto.CurrentCompareSymbol.Equals(CompareSymbols.NEQ))
                {
                    IsPartiallyEquivalent = true;
                }
                else
                {
                    var compareOp = CompareOps[GetInverse(dto.CurrentCompareSymbol)];
                    IsPartiallyEquivalent = compareOp(dto.CurrentValue, dto.PriorValue);
                }
            }

            if(dto.CurrentCompareSymbol.Equals(GetInverse(dto.PriorCompareSymbol))
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
