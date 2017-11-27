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
        public static string Mismatch => "Unreachable Case Statement: Type does not match the Select Statement";
        public static string ExceedsBoundary => "Unreachable Case Statement: Invalid Value for the Select Statement Type";
        public static string CaseElse => "Unreachable Case Else Statement: All possible values are handled by prior Case statement(s)";
    }

    public sealed class SelectCaseInspection : ParseTreeInspectionBase
    {
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
            public bool IsCaseElse;
            public bool IsHandledByPriorClause;
            public bool HasInconsistentType;
            public bool HasOutofRangeValue;
            public bool MakesRemainingClausesUnreachable;

            public CaseClauseDataObject(RubberduckParserState state, string typeName, string idReferenceName, ParserRuleContext caseClause)
            {
                CaseContext = caseClause;
                RangeClauseDOs = new List<RangeClauseDataObject>();
                IsCaseElse = caseClause is VBAParser.CaseElseClauseContext;
                IsHandledByPriorClause = false;
                HasInconsistentType = false;
                HasOutofRangeValue = false;
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
            public bool IsSelectCaseBoolean;
            public bool IsStringLiteral;
            public string IdReferenceName;
            public string TypeNameNative;
            public string TypeNameTarget;
            public string CompareSymbol;
            public SelectCaseInspectionValue SingleValue;
            public SelectCaseInspectionValue MinValue;
            public SelectCaseInspectionValue MaxValue;
            public bool HasOutOfBoundsValue;
            public bool IsPreviouslyHandled;
            public bool CausesUnreachableCaseElse;

            public RangeClauseDataObject(string typeNameTarget, string idReferenceName, VBAParser.RangeClauseContext ctxt)
            {
                Context = ctxt;
                UsesIsClause = false;
                IsValueRange = false;
                IsSingleValue = true;
                IsParseable = false;
                CompareByTextOnly = false;
                MatchesSelectCaseType = true;
                IsSelectCaseBoolean = false;
                IsStringLiteral = false;
                IdReferenceName = idReferenceName;
                TypeNameNative = typeNameTarget;
                TypeNameTarget = typeNameTarget;
                CompareSymbol = CompareSymbols.EQ;
                SingleValue = null;
                MinValue = null;
                MaxValue = null;
                HasOutOfBoundsValue = false;
                IsPreviouslyHandled = false;
                CausesUnreachableCaseElse = false;
            }
        }

        public SelectCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion){ }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        private SelectStmtDataObject _selectStmtDO;
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

                _selectStmtDO = new SelectStmtDataObject(selectStmt);

                _selectStmtDO.TypeName = DetermineTheTypeName();
                if (_selectStmtDO.TypeName.Equals(string.Empty))
                {
                    continue;
                }
                else
                {
                    _selectStmtDO.SummaryClauses.IsGTMin = SelectCaseInspectionValue.CreateUpperBound(_selectStmtDO.TypeName);
                    _selectStmtDO.SummaryClauses.IsLTMax = SelectCaseInspectionValue.CreateLowerBound(_selectStmtDO.TypeName);
                }

                var caseClauseCtxts = ParserRuleContextHelper.GetChildren<VBAParser.CaseClauseContext>(_selectStmtDO.QualifiedCtxt.Context);
                var idRefName = _selectStmtDO.IdReference != null ? _selectStmtDO.IdReference.IdentifierName : string.Empty;
                _selectStmtDO.CaseClauseDOs = caseClauseCtxts.Select(cc => CreateCaseClauseDataObject(_selectStmtDO.TypeName, idRefName, cc)).ToList();

                var reportableCaseClauseResults = EvaluateSelectStmtCaseClauses();

                foreach (var clauseDTO in reportableCaseClauseResults)
                {
                    string msg = string.Empty;
                    if (clauseDTO.IsHandledByPriorClause)
                    {
                        msg = CaseInspectionMessages.Unreachable;
                    }
                    else if (clauseDTO.HasInconsistentType)
                    {
                        msg = CaseInspectionMessages.Mismatch;
                    }
                    else if (clauseDTO.HasOutofRangeValue)
                    {
                        msg = CaseInspectionMessages.ExceedsBoundary;
                    }
                    else if (clauseDTO.IsCaseElse && _selectStmtDO.HasUnreachableCaseElse)
                    {
                        msg = CaseInspectionMessages.CaseElse;
                    }

                    if (!msg.Equals(string.Empty))
                    {
                        inspResults.Add(CreateInspectionResult(selectStmt, clauseDTO.CaseContext, msg));
                    }
                }
            }
            return inspResults;
        }

        private CaseClauseDataObject CreateCaseClauseDataObject(string typeName, string idRefName, VBAParser.CaseClauseContext ctxt)
        {
            var ccDO = new CaseClauseDataObject(State, _selectStmtDO.TypeName, idRefName, ctxt);
            var rangeClauseContexts = ParserRuleContextHelper.GetChildren<VBAParser.RangeClauseContext>(ctxt);
            foreach (var rangeClauseCtxt in rangeClauseContexts)
            {
                var rcDO = new RangeClauseDataObject(typeName, idRefName, rangeClauseCtxt);
                rcDO = InitializeRangeClauseDataObject(rcDO);
                ccDO.HasInconsistentType = !rcDO.IsParseable && !rcDO.MatchesSelectCaseType;
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

        private string GetTextForRelationalOpContext(ref RangeClauseDataObject rangeClauseDO,  VBAParser.RelationalOpContext relationalOpCtxt)
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

        private string DetermineTheTypeName()
        {
            var selectExprCtxt = ParserRuleContextHelper.GetChild<VBAParser.SelectExpressionContext>(_selectStmtDO.QualifiedCtxt.Context);

            var relationalOpCtxt = ParserRuleContextHelper.GetChild<VBAParser.RelationalOpContext>(selectExprCtxt);
            if (relationalOpCtxt != null)
            {
                return Tokens.Boolean;
            }
            else
            {
                var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(selectExprCtxt);
                if (smplName != null)
                {
                     _selectStmtDO.IdReference = GetTheSelectCaseReference(_selectStmtDO.QualifiedCtxt, smplName.GetText());
                    if (_selectStmtDO.IdReference != null)
                    {
                        return _selectStmtDO.IdReference.Declaration.AsTypeName;
                    }
                }
            }
            return string.Empty;
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

        private List<CaseClauseDataObject> EvaluateSelectStmtCaseClauses()
        {
            CheckBoundaries(_selectStmtDO.SummaryClauses);

            var reportableCaseClauseResults = new List<CaseClauseDataObject>();
            for (var idx = 0; idx < _selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = _selectStmtDO.CaseClauseDOs[idx];
                caseClause.IsHandledByPriorClause = _selectStmtDO.HasUnreachableCaseElse;

                if (caseClause.HasInconsistentType
                    || caseClause.HasOutofRangeValue
                    || caseClause.IsHandledByPriorClause)
                {
                    reportableCaseClauseResults.Add(caseClause);
                    continue;
                }
                else
                {
                    _selectStmtDO.SummaryClauses = EvaluateCaseClause(ref caseClause, _selectStmtDO.SummaryClauses);
                    if (caseClause.IsHandledByPriorClause)
                    {
                        reportableCaseClauseResults.Add(caseClause);
                    }
                    if (!_selectStmtDO.HasUnreachableCaseElse)
                    {
                        _selectStmtDO.HasUnreachableCaseElse = caseClause.MakesRemainingClausesUnreachable;
                    }
                }
            }
            if (_selectStmtDO.HasUnreachableCaseElse)
            {
                var idRefName = _selectStmtDO.IdReference != null ? _selectStmtDO.IdReference.IdentifierName : string.Empty;
                reportableCaseClauseResults.Add(new CaseClauseDataObject(State, _selectStmtDO.TypeName, idRefName, _selectStmtDO.CaseElseContext));
            }
            return reportableCaseClauseResults;
        }

        private SummaryCaseCoverage EvaluateCaseClause(ref CaseClauseDataObject caseClause, SummaryCaseCoverage summaryClauses)
        {
            for (var idx = 0; idx < caseClause.RangeClauseDOs.Count(); idx++ )
            {
                var range = caseClause.RangeClauseDOs[idx];
                if (range.IsSingleValue)
                {
                    if (range.CompareByTextOnly)
                    {
                        if (summaryClauses.Indeterminants.Contains(range.Context.GetText()))
                        {
                            range.IsPreviouslyHandled = true;
                        }
                        else
                        {
                            summaryClauses.Indeterminants.Add(range.Context.GetText());
                        }
                        caseClause.RangeClauseDOs[idx] = range;
                        continue;
                    }

                    if (!range.UsesIsClause)
                    {
                        summaryClauses = HandleSimpleSingleValueCompare(ref range, range.SingleValue, summaryClauses);
                    }
                    else  //Uses 'Is' clauses
                    {
                        if (new string[] { CompareSymbols.LT, CompareSymbols.LTE, CompareSymbols.GT, CompareSymbols.GTE }.Contains(range.CompareSymbol))
                        {
                            summaryClauses = UpdateIsLTMaxIsGTMin(range.SingleValue, range.CompareSymbol, summaryClauses);
                        }
                        else if (CompareSymbols.EQ.Equals(range.CompareSymbol))
                        {
                            summaryClauses = HandleSimpleSingleValueCompare(ref range, range.SingleValue, summaryClauses);
                        }
                        else if (CompareSymbols.NEQ.Equals(range.CompareSymbol))
                        {
                            if(summaryClauses.SingleValues.Contains(range.SingleValue))
                            {
                                range.CausesUnreachableCaseElse = true;
                            }

                            summaryClauses = UpdateIsLTMaxIsGTMin(range.SingleValue, CompareSymbols.LT, summaryClauses);
                            summaryClauses = UpdateIsLTMaxIsGTMin(range.SingleValue, CompareSymbols.GT, summaryClauses);
                        }
                    }
                }
                else  //It is a Case Statemnt like "Case 45 To 150"
                {
                    summaryClauses = AggregateRanges(summaryClauses);

                    range.IsPreviouslyHandled = summaryClauses.Ranges.Where(rg => range.MinValue.IsWithin(rg.Item1, rg.Item2) 
                            && range.MaxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                            || summaryClauses.IsLTMax != null && summaryClauses.IsLTMax > range.MaxValue
                            || summaryClauses.IsGTMin != null && summaryClauses.IsGTMin < range.MinValue;

                    if (!range.IsPreviouslyHandled)
                    {
                        var overlapsMin = summaryClauses.Ranges.Where(rg => range.MinValue.IsWithin(rg.Item1, rg.Item2));
                        var overlapsMax = summaryClauses.Ranges.Where(rg => range.MaxValue.IsWithin(rg.Item1, rg.Item2));
                        var updated = new List<Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>>();
                        foreach (var rg in summaryClauses.Ranges)
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
                        summaryClauses.Ranges = updated;
                    }

                    if (_selectStmtDO.TypeName.Equals(Tokens.Boolean))
                    {
                        range.CausesUnreachableCaseElse = range.MinValue != range.MaxValue;
                    }
                }
                caseClause.RangeClauseDOs[idx] = range;
            }

            caseClause.MakesRemainingClausesUnreachable = caseClause.RangeClauseDOs.Where(rg => rg.CausesUnreachableCaseElse).Any();

            if (!caseClause.MakesRemainingClausesUnreachable)
            {
                caseClause.MakesRemainingClausesUnreachable = IsClausesCoverAllValues(summaryClauses);
            }

            caseClause.IsHandledByPriorClause = caseClause.RangeClauseDOs.Where(rg => rg.IsPreviouslyHandled).Count() == caseClause.RangeClauseDOs.Count;
            return summaryClauses;
        }

        private bool IsClausesCoverAllValues(SummaryCaseCoverage priorHandlers)
        {
            if (priorHandlers.IsLTMax != null && priorHandlers.IsLTMax != null)
            {
                return priorHandlers.IsLTMax > priorHandlers.IsGTMin
                        || (priorHandlers.IsLTMax >= priorHandlers.IsGTMin 
                        && priorHandlers.SingleValues.Contains(priorHandlers.IsLTMax));
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

            if(CompareSymbols.LTE == compareSymbol || CompareSymbols.GTE == compareSymbol)
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
            range.IsPreviouslyHandled = SingleValueIsHandledPreviously(theValue, priorHandlers);

            if (theValue.TargetTypeName.Equals(Tokens.Boolean))
            {
                range.CausesUnreachableCaseElse = priorHandlers.SingleValues.Any()
                    && !priorHandlers.SingleValues.Contains(theValue);
            }

            if (!range.IsPreviouslyHandled)
            {
                priorHandlers.SingleValues.Add(theValue);
            }
            return priorHandlers;
        }

        private void CheckBoundaries(SummaryCaseCoverage summaryClauses)
        {
            var reportableCaseClauseResults = new List<CaseClauseDataObject>();

            if (summaryClauses.IsGTMin == null && summaryClauses.IsLTMax == null)
            {
                return;
            }

            for (var idx = 0; idx < _selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = _selectStmtDO.CaseClauseDOs[idx];
                for (var rgIdx = 0; rgIdx < caseClause.RangeClauseDOs.Count(); rgIdx++ )
                {
                    var range = caseClause.RangeClauseDOs[rgIdx];
                    if (!range.IsParseable)
                    {
                        continue;
                    }

                    if (range.IsSingleValue)
                    {
                        if (CompareSymbols.EQ.Equals(range.CompareSymbol) )
                        {
                            range.HasOutOfBoundsValue = range.SingleValue < summaryClauses.IsLTMax || range.SingleValue > summaryClauses.IsGTMin;
                        }
                        else if (CompareSymbols.GT.Equals(range.CompareSymbol) || CompareSymbols.GTE.Equals(range.CompareSymbol) )
                        {
                            range.HasOutOfBoundsValue = range.SingleValue > summaryClauses.IsGTMin;
                        }
                        else if (CompareSymbols.LT.Equals(range.CompareSymbol) || CompareSymbols.LTE.Equals(range.CompareSymbol) )
                        {
                            range.HasOutOfBoundsValue = range.SingleValue < summaryClauses.IsLTMax;
                        }
                    }
                    else
                    {
                        range.HasOutOfBoundsValue = range.MinValue > summaryClauses.IsGTMin || range.MaxValue < summaryClauses.IsLTMax;
                    }
                    caseClause.RangeClauseDOs[rgIdx] = range;
                }
                caseClause.HasOutofRangeValue = caseClause.RangeClauseDOs.Where(rg => rg.HasOutOfBoundsValue).Count() == caseClause.RangeClauseDOs.Count;
                _selectStmtDO.CaseClauseDOs[idx] = caseClause;
            }
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
                var theNextMin = ranges[idx+1].Item1;
                var theNextMax = ranges[idx+1].Item2;
                if(theMax.AsLong() == theNextMin.AsLong() - 1)
                {
                    updatedHandlers.Add(new Tuple<SelectCaseInspectionValue, SelectCaseInspectionValue>(theMin,theNextMax));
                    combinedLastRange = true;
                }
                else if(theMin.AsLong() == theNextMax.AsLong() + 1)
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
