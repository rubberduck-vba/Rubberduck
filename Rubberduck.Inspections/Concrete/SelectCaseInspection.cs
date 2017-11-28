using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    //TODO: Add replace with UI Resource
    public static class CaseInspectionMessages
    {
        public static string Unreachable => "Unreachable Case Statement: Handled by previous Case statement(s)";
        public static string MismatchType => "Unreachable Case Statement: Type does not match the Select Statement";
        public static string ExceedsBoundary => "Unreachable Case Statement: Invalid Value for the Select Statement Type";
        public static string CaseElse => "Unreachable Case Else Statement: All possible values are handled by prior Case statement(s)";
    }


    public sealed class SelectCaseInspection : ParseTreeInspectionBase
    {
        internal enum ClauseEvaluationResult { Unreachable, MismatchType, ExceedsBoundary, CaseElse, NoResult };

        internal Dictionary<ClauseEvaluationResult, string> _resultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            { ClauseEvaluationResult.Unreachable, CaseInspectionMessages.Unreachable },
            { ClauseEvaluationResult.MismatchType, CaseInspectionMessages.MismatchType },
            { ClauseEvaluationResult.ExceedsBoundary, CaseInspectionMessages.ExceedsBoundary },
            { ClauseEvaluationResult.CaseElse, CaseInspectionMessages.CaseElse }
        };

        internal static class CompareSymbols
        {
            public static readonly string EQ = "=";
            public static readonly string NEQ = "<>";
            public static readonly string LT = "<";
            public static readonly string LTE = "<=";
            public static readonly string GT = ">";
            public static readonly string GTE = ">=";
        }

        internal struct SummaryCaseCoverage
        {
            public SelectCaseInspectionValue IsLTMax;
            public SelectCaseInspectionValue IsGTMin;
            public List<SelectCaseInspectionValue> SingleValues;
            public List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>> Ranges;
            public List<string> Indeterminants;
        }

        internal struct SelectStmtDataObject
        {
            public QualifiedContext<ParserRuleContext> QualifiedCtxt;
            public string TypeName;
            public IdentifierReference IdReference;
            public List<CaseClauseDataObject> CaseClauseDOs;
            public VBAParser.CaseElseClauseContext CaseElseContext;
            public SummaryCaseCoverage SummaryClauses;
            public bool HasUnreachableCaseElse;

            public SelectStmtDataObject(QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                QualifiedCtxt = selectStmtCtxt;
                TypeName = string.Empty;
                IdReference = null;
                CaseClauseDOs = new List<CaseClauseDataObject>();
                CaseElseContext = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(QualifiedCtxt.Context);
                HasUnreachableCaseElse = false;
                SummaryClauses = new SummaryCaseCoverage
                {
                    IsGTMin = null,
                    IsLTMax = null,
                    SingleValues = new List<SelectCaseInspectionValue>(),
                    Ranges = new List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>>(),
                    Indeterminants = new List<string>()
                };
            }
        }

        internal struct CaseClauseDataObject
        {
            public ParserRuleContext CaseContext;
            public List<RangeClauseDataObject> RangeClauseDOs;
            public ClauseEvaluationResult ResultType;
            public bool MakesRemainingClausesUnreachable;

            public CaseClauseDataObject(ParserRuleContext caseClause)
            {
                CaseContext = caseClause;
                RangeClauseDOs = new List<RangeClauseDataObject>();
                ResultType = ClauseEvaluationResult.NoResult;
                MakesRemainingClausesUnreachable = false;
            }
        }

        internal struct RangeClauseDataObject
        {
            public VBAParser.RangeClauseContext Context;
            public bool UsesIsClause;
            public bool IsValueRange;
            public bool IsSingleValue;
            public bool IsParseable;
            public bool CompareByTextOnly;
            public bool MatchesSelectCaseType;
            public string IdReferenceName;
            public string TypeNameNative;
            public string TypeNameTarget;
            public string CompareSymbol;
            public SelectCaseInspectionValue SingleValue;
            public SelectCaseInspectionValue MinValue;
            public SelectCaseInspectionValue MaxValue;
            public ClauseEvaluationResult EvaluationResult;

            public RangeClauseDataObject(string typeNameTarget, string idReferenceName, VBAParser.RangeClauseContext ctxt)
            {
                Context = ctxt;
                UsesIsClause = false;
                IsValueRange = false;
                IsSingleValue = true;
                IsParseable = false;
                CompareByTextOnly = false;
                MatchesSelectCaseType = true;
                IdReferenceName = idReferenceName;
                TypeNameNative = typeNameTarget;
                TypeNameTarget = typeNameTarget;
                CompareSymbol = CompareSymbols.EQ;
                SingleValue = null;
                MinValue = null;
                MaxValue = null;
                EvaluationResult = ClauseEvaluationResult.NoResult;
            }
        }

        public SelectCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        private bool IsParseableSelectStmt(SelectStmtDataObject sdo) => !sdo.TypeName.Equals(string.Empty);

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

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            var inspResults = new List<IInspectionResult>();

            foreach (var selectStmt in selectCaseContexts)
            {
                var sdo = new SelectStmtDataObject(selectStmt);
                sdo = InitializeSelectStatementDataObject(sdo);
                if (IsParseableSelectStmt(sdo))
                {
                    sdo = InspectSelectStmtCaseClauses(sdo);
                    inspResults.AddRange(sdo.CaseClauseDOs.Where(cc => cc.ResultType != ClauseEvaluationResult.NoResult)
                        .Select(cc => CreateInspectionResult(selectStmt, cc.CaseContext, _resultMessages[cc.ResultType])));

                    var inspectedCaseClauseDOs = sdo.CaseClauseDOs;
                    if (sdo.HasUnreachableCaseElse && sdo.CaseElseContext != null)
                    {
                        inspResults.Add(CreateInspectionResult(selectStmt, sdo.CaseElseContext, _resultMessages[ClauseEvaluationResult.CaseElse]));
                    }
                }
            }
            return inspResults;
        }

        private SelectStmtDataObject InitializeSelectStatementDataObject(SelectStmtDataObject sdo)
        {
            sdo = DetermineTheTypeName(sdo);
            if (IsParseableSelectStmt(sdo))
            {
                sdo.SummaryClauses.IsGTMin = SelectCaseInspectionValue.CreateUpperBound(sdo.TypeName);
                sdo.SummaryClauses.IsLTMax = SelectCaseInspectionValue.CreateLowerBound(sdo.TypeName);
                var caseClauseCtxts = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(sdo.QualifiedCtxt.Context);
                var idRefName = sdo.IdReference != null ? sdo.IdReference.IdentifierName : string.Empty;
                sdo.CaseClauseDOs = caseClauseCtxts.Select(cc => CreateCaseClauseDataObject(sdo.TypeName, idRefName, cc)).ToList();
            }
            return sdo;
        }

        private CaseClauseDataObject CreateCaseClauseDataObject(string typeName, string idRefName, VBAParser.CaseClauseContext ctxt)
        {
            var ccDO = new CaseClauseDataObject(ctxt);
            var rangeClauseContexts = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(ctxt);
            foreach (var rangeClauseCtxt in rangeClauseContexts)
            {
                var rcDO = new RangeClauseDataObject(typeName, idRefName, rangeClauseCtxt);
                rcDO = InitializeRangeClauseDataObject(rcDO);
                ccDO.ResultType = !rcDO.IsParseable && !rcDO.MatchesSelectCaseType ? ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;
                ccDO.RangeClauseDOs.Add(rcDO);
            }
            return ccDO;
        }

        private RangeClauseDataObject InitializeRangeClauseDataObject(RangeClauseDataObject rangeClauseDO)
        {
            rangeClauseDO.UsesIsClause = HasChildToken(rangeClauseDO.Context, Tokens.Is);
            rangeClauseDO.IsValueRange = HasChildToken(rangeClauseDO.Context, Tokens.To);
            rangeClauseDO = SetTheCompareOperator(rangeClauseDO);
            rangeClauseDO.IsSingleValue = !rangeClauseDO.IsValueRange;
            rangeClauseDO.TypeNameNative = EvaluateRangeClauseTypeName(rangeClauseDO);

            if (rangeClauseDO.IsValueRange)
            {
                var startValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(rangeClauseDO.Context));
                var endValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(rangeClauseDO.Context));

                var startValue = new SelectCaseInspectionValue(startValueAsString, rangeClauseDO.TypeNameTarget);
                var endValue = new SelectCaseInspectionValue(endValueAsString, rangeClauseDO.TypeNameTarget);

                rangeClauseDO.MinValue = startValue <= endValue ? startValue : endValue;
                rangeClauseDO.MaxValue = startValue <= endValue ? endValue : startValue;
                rangeClauseDO.SingleValue = rangeClauseDO.MinValue;
                rangeClauseDO.IsParseable = rangeClauseDO.MinValue.IsParseable && rangeClauseDO.MaxValue.IsParseable;
            }
            else
            {
                rangeClauseDO.SingleValue = new SelectCaseInspectionValue(GetRangeClauseText(ref rangeClauseDO), rangeClauseDO.TypeNameTarget);
                rangeClauseDO.MaxValue = rangeClauseDO.MinValue;
                rangeClauseDO.MinValue = rangeClauseDO.SingleValue;
                rangeClauseDO.IsParseable = rangeClauseDO.SingleValue.IsParseable;
            }

            rangeClauseDO.CompareByTextOnly = !rangeClauseDO.IsParseable && rangeClauseDO.TypeNameNative.Equals(rangeClauseDO.TypeNameTarget);
            rangeClauseDO.MatchesSelectCaseType = rangeClauseDO.TypeNameTarget.Equals(rangeClauseDO.TypeNameNative);

            return rangeClauseDO;
        }
        #region RangeClauseDOPort
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

        private RangeClauseDataObject SetTheCompareOperator(RangeClauseDataObject rangeClauseDO)
        {
            VBAParser.ComparisonOperatorContext opCtxt;
            rangeClauseDO.UsesIsClause = TryGetChildContext(rangeClauseDO.Context, out opCtxt);
            rangeClauseDO.CompareSymbol = rangeClauseDO.UsesIsClause ? opCtxt.GetText() : CompareSymbols.EQ;
            return rangeClauseDO;
        }

        private static bool TryGetChildContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : ParserRuleContext
        {
            opCtxt = null;
            opCtxt = ParserRuleContextHelper.GetChild<U>(ctxt);
            return opCtxt != null;
        }

        private bool IsStringLiteral(string text) => text.StartsWith("\"") && text.EndsWith("\"");

        private string EvaluateRangeClauseTypeName(RangeClauseDataObject rangClauseDO)
        {
            var textValue = rangClauseDO.Context.GetText();
            if (IsStringLiteral(textValue))
            {
                return Tokens.String;
            }
            else if (textValue.EndsWith("#"))
            {
                var modified = textValue.Substring(0, textValue.Length - 1);
                if (long.TryParse(modified, out _))
                {
                    return Tokens.Double;
                }
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
                return rangClauseDO.TypeNameTarget;
            }
            else if (textValue.Equals(Tokens.True) || textValue.Equals(Tokens.False))
            {
                return Tokens.Boolean;
            }
            else
            {
                return rangClauseDO.TypeNameTarget;
            }
        }

        private static string GetText(ParserRuleContext ctxt) => ctxt.GetText().Replace("\"", "");

        private string GetRangeClauseText(ref RangeClauseDataObject rangeClauseDO)
        {
            var ctxt = rangeClauseDO.Context;
            VBAParser.RelationalOpContext relationalOpCtxt;
            if (TryGetChildContext(ctxt, out relationalOpCtxt))
            {
                rangeClauseDO.UsesIsClause = true;
                return GetTextForRelationalOpContext(ref rangeClauseDO, relationalOpCtxt);
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

        private string GetTextForRelationalOpContext(ref RangeClauseDataObject rangeClauseDO, VBAParser.RelationalOpContext relationalOpCtxt)
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
                else if (relationalOpCtxt.children[idx] is VBAParser.UnaryMinusOpContext
                        || relationalOpCtxt.children[idx] is VBAParser.LiteralExprContext)
                {
                    literalExprCtxtIndices.Add(idx);
                }
                else if (IsComparisonOperator(text))
                {
                    rangeClauseDO.CompareSymbol = text;
                }
            }

            if (lExprCtxtIndices.Count() == 2)  //e.g., x > someConstantExpression
            {
                var ctxtLHS = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.First()];
                var ctxtRHS = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.Last()];

                string result;
                if (GetText(ctxtLHS).Equals(rangeClauseDO.IdReferenceName))
                {
                    return TryGetTheExpressionValue(ctxtRHS, out result) ? result : string.Empty;
                }
                else if (GetText(ctxtRHS).Equals(rangeClauseDO.IdReferenceName))
                {
                    rangeClauseDO.CompareSymbol = GetInverse(rangeClauseDO.CompareSymbol);
                    return TryGetTheExpressionValue(ctxtLHS, out result) ? result : string.Empty;
                }
            }
            else if (lExprCtxtIndices.Count == 1 && literalExprCtxtIndices.Count == 1) // e.g., z < 10
            {
                var lExpIndex = lExprCtxtIndices.First();
                var litExpIndex = literalExprCtxtIndices.First();
                var lExprCtxt = (VBAParser.LExprContext)relationalOpCtxt.children[lExpIndex];
                if (GetText(lExprCtxt).Equals(rangeClauseDO.IdReferenceName))
                {
                    rangeClauseDO.CompareSymbol = lExpIndex > litExpIndex ?
                        GetInverse(rangeClauseDO.CompareSymbol) : rangeClauseDO.CompareSymbol;
                    return GetText((ParserRuleContext)relationalOpCtxt.children[litExpIndex]);
                }
            }
            return string.Empty;
        }

        private static bool IsComparisonOperator(string opCandidate) => _compareInversions.Keys.Contains(opCandidate);

        private static string GetInverse(string theOperator)
        {
            return IsComparisonOperator(theOperator) ? _compareInversions[theOperator] : theOperator;
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

        private IdentifierReference GetTheRangeClauseReference(ParserRuleContext rangeClauseCtxt, string theName)
        {
            var allRefs = new List<IdentifierReference>();
            foreach (var dec in State.DeclarationFinder.MatchName(theName))
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
        #endregion

        private SelectStmtDataObject DetermineTheTypeName(SelectStmtDataObject sdo)
        {
            var selectExprCtxt = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(sdo.QualifiedCtxt.Context);

            var relationalOpCtxt = ParserRuleContextHelper.GetChild<VBAParser.RelationalOpContext>(selectExprCtxt);
            if (relationalOpCtxt != null)
            {
                sdo.TypeName = Tokens.Boolean;
                return sdo;
            }
            else
            {
                var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(selectExprCtxt);
                if (smplName != null)
                {
                    sdo.IdReference = GetTheSelectCaseReference(sdo.QualifiedCtxt, smplName.GetText());
                    if (sdo.IdReference != null)
                    {
                        sdo.TypeName = sdo.IdReference.Declaration.AsTypeName;
                        return sdo;
                    }
                }
            }
            return sdo;
        }

        private IdentifierReference GetTheSelectCaseReference(QualifiedContext<ParserRuleContext> selectCaseStmt, string theName)
        {
            var allRefs = new List<IdentifierReference>();
            foreach (var dec in State.DeclarationFinder.MatchName(theName))
            {
                allRefs.AddRange(dec.References);
            }

            var selectCaseExpr = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(selectCaseStmt.Context);
            var selectCaseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, selectCaseStmt.Context)
                                    && (ParserRuleContextHelper.HasParent(rf.Context, selectCaseExpr)));

            Debug.Assert(selectCaseReference.Count() == 1);
            return selectCaseReference.First();
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private SelectStmtDataObject InspectSelectStmtCaseClauses(SelectStmtDataObject selectStmtDO)
        {
            selectStmtDO = CheckBoundaries(selectStmtDO);

            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = selectStmtDO.CaseClauseDOs[idx];
                if (selectStmtDO.HasUnreachableCaseElse)
                {
                    caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                }
                else if (caseClause.ResultType.Equals(ClauseEvaluationResult.NoResult))
                {
                    selectStmtDO = InspectCaseClause(ref caseClause, selectStmtDO);
                    if (!selectStmtDO.HasUnreachableCaseElse)
                    {
                        selectStmtDO.HasUnreachableCaseElse = caseClause.MakesRemainingClausesUnreachable;
                    }
                }
                selectStmtDO.CaseClauseDOs[idx] = caseClause;
                selectStmtDO.HasUnreachableCaseElse = selectStmtDO.CaseClauseDOs.Where(cc => cc.MakesRemainingClausesUnreachable).Any();
            }

            return selectStmtDO;
        }

        private SelectStmtDataObject InspectRangeClause(SelectStmtDataObject sdo, ref RangeClauseDataObject range)
        {
            if (range.CompareByTextOnly)
            {
                if (sdo.SummaryClauses.Indeterminants.Contains(range.Context.GetText()))
                {
                    range.EvaluationResult = ClauseEvaluationResult.Unreachable;
                }
                else
                {
                    sdo.SummaryClauses.Indeterminants.Add(range.Context.GetText());
                }
            }
            else if (range.IsSingleValue)
            {
                if (!range.UsesIsClause)
                {
                    sdo.SummaryClauses = HandleSimpleSingleValueCompare(ref range, range.SingleValue, sdo.SummaryClauses);
                }
                else  //Uses 'Is' clauses
                {
                    if (new string[] { CompareSymbols.LT, CompareSymbols.LTE, CompareSymbols.GT, CompareSymbols.GTE }.Contains(range.CompareSymbol))
                    {
                        sdo.SummaryClauses = UpdateIsLTMaxIsGTMin(range.SingleValue, range.CompareSymbol, sdo.SummaryClauses);
                    }
                    else if (CompareSymbols.EQ.Equals(range.CompareSymbol))
                    {
                        sdo.SummaryClauses = HandleSimpleSingleValueCompare(ref range, range.SingleValue, sdo.SummaryClauses);
                    }
                    else if (CompareSymbols.NEQ.Equals(range.CompareSymbol))
                    {
                        if (sdo.SummaryClauses.SingleValues.Contains(range.SingleValue))
                        {
                            range.EvaluationResult = ClauseEvaluationResult.CaseElse;
                        }

                        sdo.SummaryClauses = UpdateIsLTMaxIsGTMin(range.SingleValue, CompareSymbols.LT, sdo.SummaryClauses);
                        sdo.SummaryClauses = UpdateIsLTMaxIsGTMin(range.SingleValue, CompareSymbols.GT, sdo.SummaryClauses);
                    }
                }
            }
            else  //It is a range of values like "Case 45 To 150"
            {
                sdo.SummaryClauses = AggregateRanges(sdo.SummaryClauses);
                var minValue = range.MinValue;
                var maxValue = range.MaxValue;
                range.EvaluationResult = sdo.SummaryClauses.Ranges.Where(rg => minValue.IsWithin(rg.Item1, rg.Item2)
                        && maxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                        || sdo.SummaryClauses.IsLTMax != null && sdo.SummaryClauses.IsLTMax > range.MaxValue
                        || sdo.SummaryClauses.IsGTMin != null && sdo.SummaryClauses.IsGTMin < range.MinValue
                        ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

                if (range.EvaluationResult == ClauseEvaluationResult.NoResult)
                {
                    var overlapsMin = sdo.SummaryClauses.Ranges.Where(rg => minValue.IsWithin(rg.Item1, rg.Item2));
                    var overlapsMax = sdo.SummaryClauses.Ranges.Where(rg => maxValue.IsWithin(rg.Item1, rg.Item2));
                    var updated = new List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>>();
                    foreach (var rg in sdo.SummaryClauses.Ranges)
                    {
                        if (overlapsMin.Contains(rg))
                        {
                            updated.Add(new Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>(rg.Item1, range.MaxValue));
                        }
                        else if (overlapsMax.Contains(rg))
                        {
                            updated.Add(new Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>(range.MinValue, rg.Item2));
                        }
                        else
                        {
                            updated.Add(rg);
                        }
                    }

                    if (!overlapsMin.Any() && !overlapsMax.Any())
                    {
                        updated.Add(new Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>(range.MinValue, range.MaxValue));
                    }
                    sdo.SummaryClauses.Ranges = updated;
                }

                if (sdo.TypeName.Equals(Tokens.Boolean))
                {
                    range.EvaluationResult = range.MinValue != range.MaxValue
                        ? ClauseEvaluationResult.CaseElse : ClauseEvaluationResult.NoResult;
                }
            }
            return sdo;
        }

        private SelectStmtDataObject InspectCaseClause(ref CaseClauseDataObject caseClause, SelectStmtDataObject sdo)
        {
            for (var idx = 0; idx < caseClause.RangeClauseDOs.Count(); idx++)
            {
                var range = caseClause.RangeClauseDOs[idx];
                sdo = InspectRangeClause(sdo, ref range);
                caseClause.RangeClauseDOs[idx] = range;
            }

            caseClause.MakesRemainingClausesUnreachable =
                    caseClause.RangeClauseDOs.Where(rg => rg.EvaluationResult == ClauseEvaluationResult.CaseElse).Any()
                    || IsClausesCoverAllValues(sdo.SummaryClauses);

            caseClause.ResultType = caseClause.RangeClauseDOs.Where(rg => rg.EvaluationResult == ClauseEvaluationResult.Unreachable).Count() == caseClause.RangeClauseDOs.Count
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            return sdo;
        }

        private bool IsClausesCoverAllValues(SummaryCaseCoverage summaryClauses)
        {
            if (summaryClauses.IsLTMax != null && summaryClauses.IsLTMax != null)
            {
                return summaryClauses.IsLTMax > summaryClauses.IsGTMin
                        || (summaryClauses.IsLTMax >= summaryClauses.IsGTMin
                        && summaryClauses.SingleValues.Contains(summaryClauses.IsLTMax));
            }
            return false;
        }

        private bool SingleValueIsHandledPreviously(SelectCaseInspectionValue theValue, SummaryCaseCoverage priorHandlers)
        {
            return priorHandlers.IsLTMax != null && theValue < priorHandlers.IsLTMax
                || priorHandlers.IsGTMin != null && theValue > priorHandlers.IsGTMin
                || priorHandlers.SingleValues.Contains(theValue)
                || priorHandlers.Ranges.Where(rg => theValue.IsWithin(rg.Item1, rg.Item2)).Any();
        }

        private SummaryCaseCoverage UpdateIsLTMaxIsGTMin(SelectCaseInspectionValue theValue, string compareSymbol, SummaryCaseCoverage priorHandlers)
        {
            if (!(new string[] { CompareSymbols.LT, CompareSymbols.LTE, CompareSymbols.GT, CompareSymbols.GTE }.Contains(compareSymbol)))
            {
                return priorHandlers;
            }

            if (new string[] { CompareSymbols.LT, CompareSymbols.LTE }.Contains(compareSymbol))
            {
                priorHandlers.IsLTMax = priorHandlers.IsLTMax == null ? theValue
                    : priorHandlers.IsLTMax < theValue ? theValue : priorHandlers.IsLTMax;
            }
            else
            {
                priorHandlers.IsGTMin = priorHandlers.IsGTMin == null ? theValue
                    : priorHandlers.IsGTMin > theValue ? theValue : priorHandlers.IsGTMin;
            }

            if (CompareSymbols.LTE == compareSymbol || CompareSymbols.GTE == compareSymbol)
            {
                if (!priorHandlers.SingleValues.Contains(theValue))
                {
                    priorHandlers.SingleValues.Add(theValue);
                }
            }
            return priorHandlers;
        }

        private SummaryCaseCoverage HandleSimpleSingleValueCompare(ref RangeClauseDataObject range, SelectCaseInspectionValue theValue, SummaryCaseCoverage priorHandlers)
        {
            range.EvaluationResult = SingleValueIsHandledPreviously(theValue, priorHandlers)
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            if (theValue.TargetTypeName.Equals(Tokens.Boolean))
            {
                range.EvaluationResult = priorHandlers.SingleValues.Any()
                    && !priorHandlers.SingleValues.Contains(theValue)
                    ? ClauseEvaluationResult.CaseElse : range.EvaluationResult;
            }

            if (range.EvaluationResult != ClauseEvaluationResult.Unreachable)
            {
                priorHandlers.SingleValues.Add(theValue);
            }
            return priorHandlers;
        }

        private SelectStmtDataObject CheckBoundaries(SelectStmtDataObject selectStmtDO)
        {
            var reportableCaseClauseResults = new List<CaseClauseDataObject>();

            if (selectStmtDO.SummaryClauses.IsGTMin == null && selectStmtDO.SummaryClauses.IsLTMax == null)
            {
                return selectStmtDO;
            }

            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = selectStmtDO.CaseClauseDOs[idx];
                if (caseClause.ResultType != ClauseEvaluationResult.NoResult)
                {
                    continue;
                }

                for (var rgIdx = 0; rgIdx < caseClause.RangeClauseDOs.Count(); rgIdx++)
                {
                    var range = caseClause.RangeClauseDOs[rgIdx];
                    if (!range.IsParseable)
                    {
                        continue;
                    }

                    if (range.IsSingleValue)
                    {
                        if (CompareSymbols.EQ.Equals(range.CompareSymbol))
                        {
                            range.EvaluationResult = range.SingleValue < selectStmtDO.SummaryClauses.IsLTMax || range.SingleValue > selectStmtDO.SummaryClauses.IsGTMin
                                ? ClauseEvaluationResult.ExceedsBoundary : ClauseEvaluationResult.NoResult;
                        }
                        else if (CompareSymbols.GT.Equals(range.CompareSymbol) || CompareSymbols.GTE.Equals(range.CompareSymbol))
                        {
                            range.EvaluationResult = range.SingleValue > selectStmtDO.SummaryClauses.IsGTMin
                                ? ClauseEvaluationResult.ExceedsBoundary : ClauseEvaluationResult.NoResult;
                        }
                        else if (CompareSymbols.LT.Equals(range.CompareSymbol) || CompareSymbols.LTE.Equals(range.CompareSymbol))
                        {
                            range.EvaluationResult = range.SingleValue < selectStmtDO.SummaryClauses.IsLTMax
                                ? ClauseEvaluationResult.ExceedsBoundary : ClauseEvaluationResult.NoResult;
                        }
                    }
                    else
                    {
                        range.EvaluationResult = range.MinValue > selectStmtDO.SummaryClauses.IsGTMin || range.MaxValue < selectStmtDO.SummaryClauses.IsLTMax
                                ? ClauseEvaluationResult.ExceedsBoundary : ClauseEvaluationResult.NoResult;
                    }
                    caseClause.RangeClauseDOs[rgIdx] = range;
                }
                caseClause.ResultType = caseClause.RangeClauseDOs.Where(rg => rg.EvaluationResult == ClauseEvaluationResult.ExceedsBoundary).Count() == caseClause.RangeClauseDOs.Count ? ClauseEvaluationResult.ExceedsBoundary : ClauseEvaluationResult.NoResult;
                selectStmtDO.CaseClauseDOs[idx] = caseClause;
            }
            return selectStmtDO;
        }

        private SummaryCaseCoverage AggregateRanges(SummaryCaseCoverage currentSummaryCaseCoverage)
        {
            var startingRangeCount = currentSummaryCaseCoverage.Ranges.Count;
            if (startingRangeCount > 1)
            {
                do
                {
                    startingRangeCount = currentSummaryCaseCoverage.Ranges.Count();
                    currentSummaryCaseCoverage.Ranges = AppendRanges(currentSummaryCaseCoverage.Ranges);
                } while (currentSummaryCaseCoverage.Ranges.Count() < startingRangeCount);
            }
            return currentSummaryCaseCoverage;
        }

        private List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>> AppendRanges(List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>> ranges)
        {
            if (ranges.Count() <= 1)
            {
                return ranges;
            }

            if (!ranges.First().Item1.IsIntegerNumber)
            {
                return ranges;
            }

            var updatedHandlers = new List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>>();
            var combinedLastRange = false;

            for (var idx = 0; idx < ranges.Count(); idx++)
            {
                if (idx + 1 >= ranges.Count())
                {
                    if (!combinedLastRange)
                    {
                        updatedHandlers.Add(ranges[idx]);
                    }
                    continue;
                }
                combinedLastRange = false;
                var theMin = ranges[idx].Item1;
                var theMax = ranges[idx].Item2;
                var theNextMin = ranges[idx + 1].Item1;
                var theNextMax = ranges[idx + 1].Item2;
                if (theMax.AsLong() == theNextMin.AsLong() - 1)
                {
                    updatedHandlers.Add(new Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>(theMin, theNextMax));
                    combinedLastRange = true;
                }
                else if (theMin.AsLong() == theNextMax.AsLong() + 1)
                {
                    updatedHandlers.Add(new Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>(theNextMin, theMax));
                    combinedLastRange = true;
                }
                else
                {
                    updatedHandlers.Add(ranges[idx]);
                }
            }
            return updatedHandlers;
        }

        #region Listener
        public class UnreachableCaseInspectionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
        #endregion
    }
}