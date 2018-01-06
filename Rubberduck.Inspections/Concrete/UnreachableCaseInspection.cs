using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
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
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        internal enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse, NoResult };

        internal Dictionary<ClauseEvaluationResult, string> ResultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            [ClauseEvaluationResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
            [ClauseEvaluationResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
            [ClauseEvaluationResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        };

        internal static class CompareTokens
        {
            public static readonly string EQ = "=";
            public static readonly string NEQ = "<>";
            public static readonly string LT = "<";
            public static readonly string LTE = "<=";
            public static readonly string GT = ">";
            public static readonly string GTE = ">=";
        }

        private static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> MathOperations = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
        {
            ["*"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS * RHS; },
            ["/"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS / RHS; },
            ["+"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS + RHS; },
            ["-"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS - RHS; },
            ["^"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS ^ RHS; },
            ["Mod"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS % RHS; }
        };

        private static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> CompareOperations = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
        {
            [CompareTokens.EQ] = delegate(UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS){ return new UnreachableCaseInspectionValue(LHS == RHS ? Tokens.True: Tokens.False, Tokens.Boolean); },
            [CompareTokens.NEQ] = delegate(UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS){ return new UnreachableCaseInspectionValue(LHS != RHS ? Tokens.True: Tokens.False, Tokens.Boolean); },
            [CompareTokens.LT] = delegate(UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS){ return new UnreachableCaseInspectionValue(LHS < RHS ? Tokens.True: Tokens.False, Tokens.Boolean); },
            [CompareTokens.LTE] = delegate(UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS){ return new UnreachableCaseInspectionValue(LHS <= RHS ? Tokens.True: Tokens.False, Tokens.Boolean); },
            [CompareTokens.GT] = delegate(UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS){ return new UnreachableCaseInspectionValue(LHS > RHS ? Tokens.True: Tokens.False, Tokens.Boolean); },
            [CompareTokens.GTE] = delegate(UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS){ return new UnreachableCaseInspectionValue(LHS >= RHS ? Tokens.True: Tokens.False, Tokens.Boolean); }
        };

        internal struct SummaryCaseCoverage
        {
            public UnreachableCaseInspectionValue IsLT;
            public UnreachableCaseInspectionValue IsGT;
            public List<UnreachableCaseInspectionValue> SingleValues;
            public List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> Ranges;
            public bool CaseElseIsUnreachable;
            public List<string> RangeClausesAsText;
        }

        internal struct ExpressionEvaluationDataObject
        {
            public ParserRuleContext ParentCtxt;
            public bool IsUnaryOperation;
            public UnreachableCaseInspectionValue LHSValue;
            public UnreachableCaseInspectionValue RHSValue;
            public string Operator;
            public string SelectCaseRefName;
            public string TypeNameTarget;
            public UnreachableCaseInspectionValue Result;
            public bool CanBeInspected;
            public bool EvaluateAsIsClause;
        }

        internal struct SelectStmtDataObject
        {
            public SelectCaseStmtContext SelectStmtContext;
            public SelectExpressionContext SelectExpressionContext;
            public string BaseTypeName;
            public string AsTypeName;
            public string IdReferenceName;
            public List<CaseClauseDataObject> CaseClauseDOs;
            public CaseElseClauseContext CaseElseContext;
            public SummaryCaseCoverage SummaryCaseClauses;
            public bool CanBeInspected;

            public SelectStmtDataObject(QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                SelectStmtContext = (SelectCaseStmtContext)selectStmtCtxt.Context;
                IdReferenceName = string.Empty;
                BaseTypeName = Tokens.Variant;
                AsTypeName = Tokens.Variant;
                CaseClauseDOs = new List<CaseClauseDataObject>();
                CaseElseContext = SelectStmtContext.FindChild<CaseElseClauseContext>();
                SummaryCaseClauses = new SummaryCaseCoverage
                {
                    IsGT = null,
                    IsLT = null,
                    SingleValues = new List<UnreachableCaseInspectionValue>(),
                    Ranges = new List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>(),
                    RangeClausesAsText = new List<string>(),
                };
                CanBeInspected = TryGetChildContext(SelectStmtContext, out SelectExpressionContext);
            }
        }

        internal struct CaseClauseDataObject
        {
            public ParserRuleContext CaseContext;
            public List<RangeClauseDataObject> RangeClauseDOs;
            public ClauseEvaluationResult ResultType;

            public CaseClauseDataObject(ParserRuleContext caseClause)
            {
                CaseContext = caseClause;
                RangeClauseDOs = new List<RangeClauseDataObject>();
                ResultType = ClauseEvaluationResult.NoResult;
            }
        }

        internal struct RangeClauseDataObject
        {
            public RangeClauseContext Context;
            public bool UsesIsClause;
            public bool IsValueRange;
            public bool IsConstant;
            public bool CompareByTextOnly;
            public string IdReferenceName;
            public string AsText;
            public string TypeNameTarget;
            public string CompareSymbol;
            public UnreachableCaseInspectionValue SingleValue;
            public UnreachableCaseInspectionValue MinValue;
            public UnreachableCaseInspectionValue MaxValue;
            public ClauseEvaluationResult ResultType;
            public bool CanBeInspected;

            public RangeClauseDataObject(RangeClauseContext ctxt, string targetTypeName)
            {
                Context = ctxt;
                UsesIsClause = false;
                IsValueRange = false;
                IsConstant = false;
                CanBeInspected = true;
                CompareByTextOnly = false;
                IdReferenceName = string.Empty;
                AsText = ctxt.GetText();
                TypeNameTarget = targetTypeName;
                CompareSymbol = CompareTokens.EQ;
                SingleValue = null;
                MinValue = null;
                MaxValue = null;
                ResultType = ClauseEvaluationResult.NoResult;
            }
        }

        public UnreachableCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        //Used to modify logic operators to inspect expressions like '5 > x' as 'x < 5'
        private static Dictionary<string, string> AlgebraicLogicalInversions = new Dictionary<string, string>()
        {
            [CompareTokens.EQ] = CompareTokens.EQ,
            [CompareTokens.NEQ] = CompareTokens.NEQ,
            [CompareTokens.LT] = CompareTokens.GT,
            [CompareTokens.LTE] = CompareTokens.GTE,
            [CompareTokens.GT] = CompareTokens.LT,
            [CompareTokens.GTE] = CompareTokens.LTE
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var inspResults = new List<IInspectionResult>();

            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            foreach (var selectStmt in selectCaseContexts)
            {
                var selectStmtDO = InitializeSelectStatementDataObject(new SelectStmtDataObject(selectStmt));

                if (!selectStmtDO.CanBeInspected) { continue; }

                selectStmtDO = InitializeCaseClauses(selectStmtDO);

                if (!selectStmtDO.CanBeInspected) { continue; }

                selectStmtDO = InspectSelectStmtCaseClauses(selectStmtDO);

                if (!selectStmtDO.CanBeInspected) { continue; }

                inspResults.AddRange(selectStmtDO.CaseClauseDOs.Where(cc => cc.ResultType != ClauseEvaluationResult.NoResult)
                    .Select(cc => CreateInspectionResult(selectStmt, cc.CaseContext, ResultMessages[cc.ResultType])));

                if (selectStmtDO.SummaryCaseClauses.CaseElseIsUnreachable && selectStmtDO.CaseElseContext != null)
                {
                    inspResults.Add(CreateInspectionResult(selectStmt, selectStmtDO.CaseElseContext, ResultMessages[ClauseEvaluationResult.CaseElse]));
                }
            }
            return inspResults;
        }

        private SelectStmtDataObject InitializeSelectStatementDataObject(SelectStmtDataObject selectStmtDO)
        {
            selectStmtDO = ResolveSelectStmtInspectionType(selectStmtDO);

            if (selectStmtDO.CanBeInspected)
            {
                selectStmtDO.CaseClauseDOs = selectStmtDO.SelectStmtContext.GetChildren<CaseClauseContext>()
                    .Select(cc => CreateCaseClauseDataObject(cc, selectStmtDO.BaseTypeName)).ToList();
            }
            return selectStmtDO;
        }

        private SelectStmtDataObject ResolveSelectStmtInspectionType(SelectStmtDataObject selectStmtDO)
        {
            if (!ContextCanBeEvaluated(selectStmtDO.SelectExpressionContext, selectStmtDO.IdReferenceName))
            {
                return InferTheSelectStmtType(selectStmtDO);
            }

            if (selectStmtDO.SelectExpressionContext.GetDescendents()
                .Any(desc => IsBinaryLogicalOperation(desc) || IsUnaryLogicalOperator(desc)))
            {
                return SetTheTypeNames(selectStmtDO, Tokens.Boolean);
            }

            var firstLExpr = selectStmtDO.SelectExpressionContext.GetDescendent<LExprContext>();
            if (firstLExpr == null)
            {
                return InferTheSelectStmtType(selectStmtDO);
            }

            var expression = firstLExpr.GetDescendent<SimpleNameExprContext>().GetText();
            if (SymbolList.TypeHintToTypeName.ContainsKey(expression.Last().ToString()))
            {
                return SetTheTypeNames(selectStmtDO, SymbolList.TypeHintToTypeName[expression.Last().ToString()]);
            }

            var idRefs = (State.DeclarationFinder.MatchName(expression).Select(dec => dec.References))
                .SelectMany(rf => rf).Where(idr => idr.Context.HasParent(selectStmtDO.SelectExpressionContext));

            if (idRefs.Count() == 1)
            {
                selectStmtDO.IdReferenceName = idRefs.First().IdentifierName;
                selectStmtDO = SetTheTypeNames(selectStmtDO, idRefs.First().Declaration.AsTypeName, GetBaseTypeForDeclaration(idRefs.First().Declaration));

                if (selectStmtDO.BaseTypeName.Equals(Tokens.Variant))
                {
                    return InferTheSelectStmtType(selectStmtDO);
                }
                return selectStmtDO;
            }
            return InferTheSelectStmtType(selectStmtDO);
        }

        private SelectStmtDataObject InferTheSelectStmtType(SelectStmtDataObject selectStmtDO)
        {
            if (TryInferTypeFromRangeClauseContent(selectStmtDO, out string typeName))
            {
                return SetTheTypeNames(selectStmtDO, typeName);
            }
            selectStmtDO.CanBeInspected = false;
            return selectStmtDO;
        }

        private SelectStmtDataObject InitializeCaseClauses(SelectStmtDataObject selectStmtDO)
        {
            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count; idx++)
            {
                var caseClauseDO = selectStmtDO.CaseClauseDOs[idx];
                if (caseClauseDO.ResultType != ClauseEvaluationResult.Unreachable)
                {
                    for (var rgIdx = 0; rgIdx < caseClauseDO.RangeClauseDOs.Count; rgIdx++)
                    {
                        var rgClause = caseClauseDO.RangeClauseDOs[rgIdx];
                        rgClause.CanBeInspected = ContextCanBeEvaluated(rgClause.Context, selectStmtDO.IdReferenceName);
                        if (rgClause.CanBeInspected)
                        {
                            rgClause = InitializeRangeClauseDataObject(rgClause, selectStmtDO.BaseTypeName, selectStmtDO.IdReferenceName);
                        }
                        caseClauseDO.RangeClauseDOs[rgIdx] = rgClause;
                    }
                }
                selectStmtDO.CaseClauseDOs[idx] = caseClauseDO;
            }
            return selectStmtDO;
        }

        private SelectStmtDataObject InspectSelectStmtCaseClauses(SelectStmtDataObject selectStmtDO)
        {
            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = selectStmtDO.CaseClauseDOs[idx];

                if (selectStmtDO.SummaryCaseClauses.CaseElseIsUnreachable
                    || caseClause.RangeClauseDOs.All(rg => selectStmtDO.SummaryCaseClauses.RangeClausesAsText.Contains(rg.AsText)))
                {
                    caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                }
                else
                {
                    caseClause = InspectCaseClause(caseClause, selectStmtDO.SummaryCaseClauses);
                    selectStmtDO = UpdateSummaryClauses(selectStmtDO, caseClause);
                    if (caseClause.RangeClauseDOs.All(rg => rg.ResultType != ClauseEvaluationResult.NoResult))
                    {
                        if (caseClause.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.Unreachable))
                        {
                            caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                        }
                        else if (caseClause.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.MismatchType))
                        {
                            caseClause.ResultType = ClauseEvaluationResult.MismatchType;
                        }
                    }
                }
                selectStmtDO.CaseClauseDOs[idx] = caseClause;
            }
            return selectStmtDO;
        }

        private CaseClauseDataObject CreateCaseClauseDataObject(CaseClauseContext ctxt, string targetTypeName)
        {
            var caseClauseDO = new CaseClauseDataObject(ctxt);
            var rangeClauseContexts = ctxt.GetChildren<RangeClauseContext>();
            foreach (var rangeClauseCtxt in rangeClauseContexts)
            {
                var rgC = new RangeClauseDataObject(rangeClauseCtxt, targetTypeName);
                caseClauseDO.RangeClauseDOs.Add(rgC);
            }
            return caseClauseDO;
        }

        private RangeClauseDataObject InitializeRangeClauseDataObject(RangeClauseDataObject rangeClauseDO, string targetTypeName, string refName)
        {
            rangeClauseDO.TypeNameTarget = targetTypeName;
            rangeClauseDO.IdReferenceName = refName;
            rangeClauseDO.UsesIsClause = rangeClauseDO.Context.HasChildToken(Tokens.Is);
            rangeClauseDO.IsValueRange = rangeClauseDO.Context.HasChildToken(Tokens.To);
            rangeClauseDO = SetTheCompareOperator(rangeClauseDO);

            UnreachableCaseInspectionValue startValue;
            UnreachableCaseInspectionValue endValue;
            if (rangeClauseDO.IsValueRange)
            {
                var startContext = rangeClauseDO.Context.FindChild<SelectStartValueContext>();
                var endContext = rangeClauseDO.Context.FindChild<SelectEndValueContext>();
                rangeClauseDO = ResolveRangeClauseValue(rangeClauseDO, startContext, out startValue);
                rangeClauseDO = ResolveRangeClauseValue(rangeClauseDO, endContext, out endValue);
            }
            else
            {
                rangeClauseDO = ResolveRangeClauseValue(rangeClauseDO, rangeClauseDO.Context, out startValue);
                endValue = startValue;
            }

            if (startValue == null || !startValue.HasValue || endValue == null || !endValue.HasValue)
            {
                if ((startValue != null && !startValue.HasValue) || (endValue != null && !endValue.HasValue))
                {
                    return SetRangeClauseForTextOnlyCompare(rangeClauseDO, ClauseEvaluationResult.MismatchType);
                }
                return SetRangeClauseForTextOnlyCompare(rangeClauseDO);
            }

            if (startValue != null && endValue != null)
            {
                rangeClauseDO.MinValue = startValue <= endValue ? startValue : endValue;
                rangeClauseDO.MaxValue = startValue <= endValue ? endValue : startValue;
                rangeClauseDO.SingleValue = rangeClauseDO.MinValue;
            }
            else
            {
                return SetRangeClauseForTextOnlyCompare(rangeClauseDO, ClauseEvaluationResult.MismatchType);
            }
            return rangeClauseDO;
        }

        private RangeClauseDataObject SetRangeClauseForTextOnlyCompare(RangeClauseDataObject rangeClauseDO, ClauseEvaluationResult result = ClauseEvaluationResult.NoResult)
        {
            rangeClauseDO.CompareByTextOnly = result == ClauseEvaluationResult.NoResult;
            rangeClauseDO.CanBeInspected = false;
            rangeClauseDO.ResultType = result;
            return rangeClauseDO;
        }

        private RangeClauseDataObject ResolveRangeClauseValue(RangeClauseDataObject rangeClauseDO, ParserRuleContext context, out UnreachableCaseInspectionValue vbaValue)
        {
            vbaValue = null;
            if (!(context is RangeClauseContext || context is SelectStartValueContext || context is SelectEndValueContext))
            {
                return rangeClauseDO;
            }

            var parentEval = new ExpressionEvaluationDataObject
            {
                IsUnaryOperation = true,
                Operator = rangeClauseDO.CompareSymbol,
                CanBeInspected = rangeClauseDO.CanBeInspected,
                TypeNameTarget = rangeClauseDO.TypeNameTarget,
                SelectCaseRefName = rangeClauseDO.IdReferenceName
            };

            var contextEvals = AddEvaluationData(new Dictionary<IParseTree, ExpressionEvaluationDataObject>(), context, parentEval);
            contextEvals = ResolveContextValue(contextEvals, context);
            rangeClauseDO.CompareSymbol = contextEvals[context].Operator;
            rangeClauseDO.UsesIsClause = rangeClauseDO.UsesIsClause ? rangeClauseDO.UsesIsClause : contextEvals[context].EvaluateAsIsClause;
            vbaValue =  contextEvals[context].Result;
            return rangeClauseDO;
        }

        private Dictionary<IParseTree, ExpressionEvaluationDataObject> ResolveContextValue(Dictionary<IParseTree, ExpressionEvaluationDataObject> contextEvals, ParserRuleContext parentContext)
        {
            var parentData = GetEvaluationData(parentContext, contextEvals);
            foreach (var child in parentContext.children.Where(ch => !(ch is WhiteSpaceContext)))
            {
                var childData = GetEvaluationData(child, contextEvals);
                childData.ParentCtxt = parentContext;
                childData.TypeNameTarget = parentData.TypeNameTarget;
                childData.SelectCaseRefName = parentData.SelectCaseRefName;

                if (IsBinaryOperatorContext(child) || IsUnaryOperandContext(child))
                {
                    childData.IsUnaryOperation = IsUnaryOperandContext(child);

                    if (!childData.EvaluateAsIsClause)
                    {
                        childData.EvaluateAsIsClause = IsBinaryLogicalOperation(child) || IsUnaryLogicalOperator(child);
                    }

                    contextEvals = AddEvaluationData(contextEvals, child, childData);
                    contextEvals = ResolveContextValue(contextEvals, (ParserRuleContext)child);
                }
                else if (child is LiteralExprContext || child is LExprContext)
                {
                    childData.IsUnaryOperation = true;
                    childData.LHSValue = CreateValue((ExpressionContext)child, childData.TypeNameTarget);
                    childData.Result = childData.LHSValue;

                    contextEvals = AddEvaluationData(contextEvals, (ParserRuleContext)child, childData);
                }
                else
                {
                    contextEvals = AddEvaluationData(contextEvals, child, childData);
                }
                contextEvals = UpdateParentEvaluation(child, contextEvals);
            }
            return contextEvals;
        }

        private Dictionary<IParseTree, ExpressionEvaluationDataObject> UpdateParentEvaluation(IParseTree child, Dictionary<IParseTree, ExpressionEvaluationDataObject> ctxtEvalResults)
        {
            if (child is TerminalNodeImpl)
            {
                var terminalNode = child as TerminalNodeImpl;
                if (MathOperations.Keys.Contains(terminalNode.GetText()) || CompareOperations.Keys.Contains(terminalNode.GetText()))
                {
                    var theParentData = GetEvaluationData(terminalNode.Parent, ctxtEvalResults);
                    if (!theParentData.EvaluateAsIsClause)
                    {
                        theParentData.EvaluateAsIsClause = CompareOperations.Keys.Contains(terminalNode.GetText());
                    }
                    theParentData.Operator = terminalNode.GetText();
                    return AddEvaluationData(ctxtEvalResults, terminalNode.Parent, theParentData);
                }
                return ctxtEvalResults;
            }
            else if (child is ParserRuleContext)
            {
                return UpdateParentEvaluation((ParserRuleContext)child, ctxtEvalResults);
            }
            return ctxtEvalResults;
        }

        private Dictionary<IParseTree, ExpressionEvaluationDataObject> UpdateParentEvaluation(ParserRuleContext child, Dictionary<IParseTree, ExpressionEvaluationDataObject> ctxtEvalResults)
        {
            var childCtxt = child as ParserRuleContext;
            var childData = GetEvaluationData(childCtxt, ctxtEvalResults);
            var parentData = GetEvaluationData(childData.ParentCtxt, ctxtEvalResults);

            if (!childData.CanBeInspected)
            {
                parentData.CanBeInspected = false;
                return AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
            }

            parentData.EvaluateAsIsClause = parentData.EvaluateAsIsClause ? true : childData.EvaluateAsIsClause;
            parentData.Operator = CompareOperations.ContainsKey(childData.Operator) ? childData.Operator : parentData.Operator;

            if (parentData.IsUnaryOperation)
            {
                parentData.LHSValue = childData.Result;
                parentData.Result = childData.ParentCtxt is UnaryMinusOpContext ?
                    parentData.LHSValue.AdditiveInverse : parentData.LHSValue;
            }
            else
            {
                if (!childData.Result.HasValue && !childData.Result.AsString().Equals(childData.SelectCaseRefName))
                {
                    parentData.CanBeInspected = false;
                    return AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
                }

                if (parentData.LHSValue == null)
                {
                    parentData.LHSValue = childData.Result;
                    return AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
                }

                parentData.RHSValue = childData.Result;

                Debug.Assert(parentData.Operator != string.Empty);

                string result = string.Empty;
                if (MathOperations.ContainsKey(parentData.Operator))
                {
                    result = MathOperations[parentData.Operator](parentData.LHSValue, parentData.RHSValue).AsString();
                }
                else if (CompareOperations.ContainsKey(parentData.Operator))
                {
                    var compOperation = CompareOperations[parentData.Operator];
                    if (parentData.LHSValue.HasValue && parentData.RHSValue.HasValue)
                    {
                        //e.g., 44 < 45
                        result = compOperation(parentData.LHSValue, parentData.RHSValue).AsString();
                    }
                    else if (!parentData.LHSValue.HasValue && parentData.LHSValue.AsString().Equals(parentData.SelectCaseRefName) && parentData.RHSValue.HasValue)
                    {
                        //e.g., x < 45 is inspected as an 'Is' statement of the form 'Is < 45' .
                        result = parentData.RHSValue.AsString();
                    }
                    else if (parentData.LHSValue.HasValue && !parentData.RHSValue.HasValue && parentData.RHSValue.AsString().Equals(parentData.SelectCaseRefName))
                    {
                        //e.g., 45 > x
                        //Perform 'algebra' to get to 'x < 45' so it
                        //can be inspected as an 'Is' statement of the form 'Is < 45' .
                        parentData.Operator = AlgebraicLogicalInversions[parentData.Operator];
                        result = parentData.LHSValue.AsString();
                    }
                }
                parentData.Result = new UnreachableCaseInspectionValue(result, childData.TypeNameTarget);
            }
            return AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
        }

        private string EvaluateContextTypeName(ExpressionContext ctxt, SelectStmtDataObject selectStmtDO)
        {
            var val = CreateValue(ctxt, selectStmtDO.BaseTypeName);
            return val.HasValue ? selectStmtDO.BaseTypeName : val.DerivedTypeName;
        }

        private UnreachableCaseInspectionValue CreateValue(ExpressionContext ctxt, string typeName = "")
        {
            if (ctxt is LExprContext)
            {
                var lexprTypeName = typeName;
                if (TryGetTheLExprValue((LExprContext)ctxt, out string lexprValue, ref lexprTypeName))
                {
                    return typeName.Length > 0 ? new UnreachableCaseInspectionValue(lexprValue, typeName) : new UnreachableCaseInspectionValue(lexprValue, lexprTypeName);
                }
                var idRefs = (State.DeclarationFinder.MatchName(ctxt.GetText()).Select(dec => dec.References)).SelectMany(rf => rf)
                    .Where(idr => idr.Context.Parent == ctxt);
                if (idRefs.Any())
                {
                    var theTypeName = GetBaseTypeForDeclaration(idRefs.First().Declaration);
                    return new UnreachableCaseInspectionValue(ctxt.GetText(), theTypeName);
                }
                return new UnreachableCaseInspectionValue(ctxt.GetText(), typeName);
            }
            else if (ctxt is LiteralExprContext)
            {
                return new UnreachableCaseInspectionValue(ctxt.GetText(), typeName);
            }
            return null;
        }

        private bool TryInferTypeFromRangeClauseContent(SelectStmtDataObject selectStmtDO, out string typeName)
        {
            typeName = selectStmtDO.BaseTypeName;

            if (!selectStmtDO.BaseTypeName.Equals(Tokens.Variant)) { return false; }

            var rangeCtxts = selectStmtDO.SelectStmtContext.GetChildren<CaseClauseContext>()
                .Select(cc => cc.GetChildren<RangeClauseContext>())
                .SelectMany(rgCtxt => rgCtxt);

            var typeNames = rangeCtxts.SelectMany(rgCtxt => rgCtxt.GetDescendents())
                .Where(desc => desc is LiteralExprContext || desc is LExprContext)
                    .Select(exprCtxt => EvaluateContextTypeName((ExpressionContext)exprCtxt, selectStmtDO))
                    .Where(tn => tn != string.Empty);

            if (typeNames.All(tn => typeNames.First().Equals(tn)))
            {
                typeName = typeNames.First();
                return true;
            }

            if (typeNames.All(tn => tn.Equals(Tokens.Long)
                    || tn.Equals(Tokens.LongLong)
                    || tn.Equals(Tokens.Integer)
                    || tn.Equals(Tokens.Byte)))
            {
                typeName = Tokens.Long;
                return true;
            }

            if (typeNames.All(tn => !(tn.Equals(Tokens.Currency) || tn.Equals(Tokens.String))))
            {
                typeName = Tokens.Double;
                return true;
            }
            return false;
        }

        private bool TryGetTheLExprValue(LExprContext ctxt, out string expressionValue, ref string typeName)
        {
            expressionValue = string.Empty;
            if (TryGetChildContext(ctxt, out MemberAccessExprContext member))
            {
                var smplNameMemberRHS = member.FindChild<UnrestrictedIdentifierContext>();
                var memberDeclarations = State.DeclarationFinder.AllUserDeclarations.Where(dec => dec.IdentifierName.Equals(smplNameMemberRHS.GetText()));

                foreach (var dec in memberDeclarations)
                {
                    if (dec.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        var theCtxt = dec.Context;
                        if (theCtxt is EnumerationStmt_ConstantContext)
                        {
                            expressionValue = GetConstantDeclarationValue(dec);
                            typeName = dec.AsTypeIsBaseType ? dec.AsTypeName : dec.AsTypeDeclaration.AsTypeName;
                            return true;
                        }
                    }
                }
                return false;
            }
            else if (TryGetChildContext(ctxt, out SimpleNameExprContext smplName))
            {
                var identifierReferences = (State.DeclarationFinder.MatchName(smplName.GetText()).Select(dec => dec.References)).SelectMany(rf => rf);

                var rangeClauseReferences = identifierReferences.Where(rf => rf.Context.HasParent(smplName)
                                        && (rf.Context.HasParent(smplName.Parent)));

                var rangeClauseIdentifierReference = rangeClauseReferences.Any() ? rangeClauseReferences.First() : null;
                if (rangeClauseIdentifierReference != null)
                {
                    if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant)
                        || rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        expressionValue = GetConstantDeclarationValue(rangeClauseIdentifierReference.Declaration);
                        typeName = rangeClauseIdentifierReference.Declaration.AsTypeName;
                        return true;
                    }
                }
            }
            return false;
        }

        private string GetConstantDeclarationValue(Declaration valueDeclaration)
        {
            var contextsOfInterest = GetRHSContexts(valueDeclaration.Context.children.ToList());
            foreach (var child in contextsOfInterest)
            {
                if (IsMathOperation(child))
                {
                    var parentData = new Dictionary<IParseTree, ExpressionEvaluationDataObject>();
                    var exprEval = new ExpressionEvaluationDataObject
                    {
                        IsUnaryOperation = IsUnaryMathOperation(child),
                        Operator = CompareTokens.EQ,
                        CanBeInspected = true,
                        TypeNameTarget = valueDeclaration.AsTypeName,
                        SelectCaseRefName = valueDeclaration.IdentifierName
                    };

                    parentData = AddEvaluationData(parentData, child, exprEval);
                    return ResolveContextValue(parentData, child).First().Value.Result.AsString();
                }

                if (child is LiteralExprContext)
                {
                    if (child.Parent is EnumerationStmt_ConstantContext)
                    {
                        return child.GetText();
                    }
                    else if (valueDeclaration is ConstantDeclaration)
                    {
                        return ((ConstantDeclaration)valueDeclaration).Expression;
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
            }
            return string.Empty;
        }

        private List<ParserRuleContext> GetRHSContexts(List<IParseTree> contexts)
        {
            var contextsOfInterest = new List<ParserRuleContext>();
            var eqIndex = contexts.FindIndex(ch => ch.GetText().Equals(CompareTokens.EQ));
            if (eqIndex == contexts.Count)
            {
                return contextsOfInterest;
            }
            for (int idx = eqIndex + 1; idx < contexts.Count(); idx++)
            {
                var childCtxt = contexts[idx];
                if (!(childCtxt is WhiteSpaceContext))
                {
                    contextsOfInterest.Add((ParserRuleContext)childCtxt);
                }
            }
            return contextsOfInterest;
        }

        private string GetBaseTypeForDeclaration(Declaration declaration)
        {
            if (!declaration.AsTypeIsBaseType)
            {
                return GetBaseTypeForDeclaration(declaration.AsTypeDeclaration);
            }
            return declaration.AsTypeName;
        }

        private IdentifierReference GetTheIdReference(RuleContext parentContext, string theName)
        {
            var identifierReferences = (State.DeclarationFinder.MatchName(theName).Select(dec => dec.References)).SelectMany(rf => rf);
            var candidateIdRefs =  identifierReferences.Where(idr => parentContext.GetChildren<ParserRuleContext>().Contains(idr.Context));
            return candidateIdRefs.Any() ? candidateIdRefs.First() : null;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private SelectStmtDataObject UpdateSummaryClauses(SelectStmtDataObject selectStmtDO, CaseClauseDataObject caseClause)
        {
            foreach (var rangeClauseDO in caseClause.RangeClauseDOs)
            {
                if (rangeClauseDO.ResultType != ClauseEvaluationResult.NoResult || !rangeClauseDO.CanBeInspected)
                {
                    continue;
                }

                selectStmtDO.SummaryCaseClauses = rangeClauseDO.IsValueRange ?
                    UpdateSummaryDataRanges(selectStmtDO.SummaryCaseClauses, rangeClauseDO)
                    : UpdateSummaryDataSingleValues(selectStmtDO.SummaryCaseClauses, rangeClauseDO);

                selectStmtDO.SummaryCaseClauses.CaseElseIsUnreachable = EvaluateCaseElseAccessibility(selectStmtDO.SummaryCaseClauses, rangeClauseDO.TypeNameTarget);
            }
            selectStmtDO.SummaryCaseClauses.RangeClausesAsText.AddRange(caseClause.RangeClauseDOs.Select(rg => rg.AsText));
            return selectStmtDO;
        }

        private CaseClauseDataObject InspectCaseClause(CaseClauseDataObject caseClause, SummaryCaseCoverage summaryCoverage)
        {
            if (caseClause.ResultType != ClauseEvaluationResult.NoResult)
            {
                return caseClause;
            }

            for (var idx = 0; idx < caseClause.RangeClauseDOs.Count(); idx++)
            {
                if (!caseClause.RangeClauseDOs[idx].CanBeInspected 
                    || caseClause.RangeClauseDOs[idx].ResultType != ClauseEvaluationResult.NoResult)
                {
                    continue;
                }

                caseClause.RangeClauseDOs[idx] = caseClause.RangeClauseDOs[idx].IsValueRange ?
                    InspectValueRangeRangeClause(summaryCoverage, caseClause.RangeClauseDOs[idx])
                    : InspectSingleValueRangeClause(summaryCoverage, caseClause.RangeClauseDOs[idx]);
            }
            return caseClause;
        }

        private RangeClauseDataObject InspectValueRangeRangeClause(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO)
        {
            if (rangeClauseDO.MinValue != null && rangeClauseDO.MaxValue != null)
            {
                var isUnreachable = summaryCoverage.Ranges.Any(rg => rangeClauseDO.MinValue.IsWithin(rg.Item1, rg.Item2)
                                   && rangeClauseDO.MaxValue.IsWithin(rg.Item1, rg.Item2))
                                    || rangeClauseDO.MinValue.ExceedsMaxMin() && rangeClauseDO.MaxValue.ExceedsMaxMin()
                                    || summaryCoverage.IsLT != null && summaryCoverage.IsLT > rangeClauseDO.MaxValue
                                    || summaryCoverage.IsGT != null && summaryCoverage.IsGT < rangeClauseDO.MinValue
                                    || CheckEverything(summaryCoverage, rangeClauseDO);

                rangeClauseDO.ResultType = isUnreachable ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
            }
            return rangeClauseDO;
        }

        private bool CheckEverything(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO)
        {
            if (IsIntegerNumberType(rangeClauseDO.TypeNameTarget))
            {
                if (summaryCoverage.IsGT != null)
                {
                    //e.g., Is > 50 Range: 30 to 75 
                    if (summaryCoverage.IsGT < rangeClauseDO.MaxValue)
                    {
                        var overlapMin = rangeClauseDO.MinValue.AsLong().Value;
                        var overlapMax = summaryCoverage.IsGT.AsLong().Value;
                        return EvaluateCoverageGaps(summaryCoverage, rangeClauseDO, overlapMin, overlapMax);
                    }
                }
                if (summaryCoverage.IsLT != null)
                {
                    //e.g., Is < 50 Range: 30 to 75
                    if (summaryCoverage.IsLT > rangeClauseDO.MinValue)
                    {
                        var overlapMin = summaryCoverage.IsLT.AsLong().Value;
                        var overlapMax = rangeClauseDO.MaxValue.AsLong().Value;
                        return EvaluateCoverageGaps(summaryCoverage, rangeClauseDO, overlapMin, overlapMax);
                    }
                }
            }
            return false;
        }

        private bool EvaluateCoverageGaps(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO, long overlapMin, long overlapMax)
        {
            var reachableValues = new List<long>();
            for (var overlapValue = overlapMin; overlapValue <= overlapMax; overlapValue++)
            {
                var evalNum = new UnreachableCaseInspectionValue(overlapValue);
                if (summaryCoverage.SingleValues.Contains(evalNum)
                    || summaryCoverage.Ranges.Any(rg => evalNum.IsWithin(rg.Item1,rg.Item2)))
                {
                    reachableValues.Add(overlapValue);
                }
            }
            return reachableValues.Count() == overlapMax - overlapMin + 1;
        }

        private RangeClauseDataObject InspectSingleValueRangeClause(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO)
        {
            if (rangeClauseDO.SingleValue == null)
            {
                return rangeClauseDO;
            }

            if (rangeClauseDO.SingleValue.ExceedsMaxMin())
            {
                rangeClauseDO.ResultType = ClauseEvaluationResult.Unreachable;
                return rangeClauseDO;
            }

            var isUnreachable = false;
            if (rangeClauseDO.UsesIsClause)
            {
                if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.LT)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.LTE))
                {
                    isUnreachable = summaryCoverage.IsLT != null && summaryCoverage.IsLT >= rangeClauseDO.SingleValue;
                }
                else if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.GT)
                        || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GTE))
                {
                        isUnreachable = summaryCoverage.IsGT != null && summaryCoverage.IsGT <= rangeClauseDO.SingleValue;
                }
                else if (CompareTokens.EQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    isUnreachable = SingleValueIsHandledPreviously(rangeClauseDO.SingleValue, summaryCoverage);
                }
                else if (CompareTokens.NEQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    if (rangeClauseDO.TypeNameTarget.Equals(Tokens.Boolean))
                    {
                        isUnreachable = (rangeClauseDO.SingleValue == UnreachableCaseInspectionValue.False ?
                            summaryCoverage.SingleValues.Any(sv => sv.AsLong().Value != 0)
                            : summaryCoverage.SingleValues.Any(sv => sv.AsLong().Value == 0));
                    }
                }
            }
            else
            {
                isUnreachable = SingleValueIsHandledPreviously(rangeClauseDO.SingleValue, summaryCoverage);
            }
            rangeClauseDO.ResultType = isUnreachable ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
            return rangeClauseDO;
        }

        private bool SingleValueIsHandledPreviously(UnreachableCaseInspectionValue theValue, SummaryCaseCoverage summaryClauses)
        {
            if (theValue.UseageTypeName.Equals(Tokens.Boolean))
            {
                return summaryClauses.SingleValues.Any(val => val.AsBoolean() == theValue.AsBoolean());
            }
            else
            {
                return summaryClauses.IsLT != null && theValue < summaryClauses.IsLT
                    || summaryClauses.IsGT != null && theValue > summaryClauses.IsGT
                    || summaryClauses.SingleValues.Contains(theValue)
                    || summaryClauses.Ranges.Any(rg => theValue.IsWithin(rg.Item1, rg.Item2));
            }
        }

        private SummaryCaseCoverage UpdateSummaryIsClauseLimits(UnreachableCaseInspectionValue theValue, string compareSymbol, SummaryCaseCoverage priorHandlers)
        {
            Debug.Assert(theValue != null);

            if (compareSymbol.Equals(CompareTokens.LT) || compareSymbol.Equals(CompareTokens.LTE))
            {
                priorHandlers.IsLT = priorHandlers.IsLT ?? theValue;
                if (priorHandlers.IsLT < theValue)
                {
                    priorHandlers.IsLT = theValue;
                }
                if (theValue.UseageTypeName.Equals(Tokens.Byte))
                {
                   priorHandlers.SingleValues = LoadRangeOfByteValues(priorHandlers.SingleValues, UnreachableCaseInspectionValue.MinValueByte, priorHandlers.IsLT.AsLong().Value - 1);
                }
            }
            else if (compareSymbol.Equals(CompareTokens.GT) || compareSymbol.Equals(CompareTokens.GTE))
            {
                priorHandlers.IsGT = priorHandlers.IsGT ?? theValue;
                if (priorHandlers.IsGT > theValue)
                {
                    priorHandlers.IsGT = theValue;
                }
                if (theValue.UseageTypeName.Equals(Tokens.Byte))
                {
                    priorHandlers.SingleValues = LoadRangeOfByteValues(priorHandlers.SingleValues, priorHandlers.IsGT.AsLong().Value + 1, UnreachableCaseInspectionValue.MaxValueByte);
                }
            }
            else
            {
                return priorHandlers;
            }

            if (CompareTokens.LTE == compareSymbol || CompareTokens.GTE == compareSymbol)
            {
                priorHandlers.SingleValues.Add(theValue);
            }
            return priorHandlers;
        }

        private SummaryCaseCoverage UpdateSummaryDataRanges(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO)
        {
            if (!rangeClauseDO.CanBeInspected || !rangeClauseDO.IsValueRange) { return summaryCoverage; }

            if (rangeClauseDO.TypeNameTarget.Equals(Tokens.Boolean))
            {
                if (rangeClauseDO.MinValue != UnreachableCaseInspectionValue.Zero || rangeClauseDO.MaxValue != UnreachableCaseInspectionValue.Zero)
                {
                    summaryCoverage.SingleValues.Add(UnreachableCaseInspectionValue.True);
                }
                if (UnreachableCaseInspectionValue.Zero.IsWithin(rangeClauseDO.MinValue, rangeClauseDO.MaxValue))
                {
                    summaryCoverage.SingleValues.Add(UnreachableCaseInspectionValue.False);
                }
            }

            if (rangeClauseDO.TypeNameTarget.Equals(Tokens.Byte))
            {
                summaryCoverage.SingleValues = LoadRangeOfByteValues(summaryCoverage.SingleValues, rangeClauseDO.MinValue.AsLong().Value, rangeClauseDO.MaxValue.AsLong().Value);
            }

            var updatedRanges = new List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>();
            var overlapsMin = summaryCoverage.Ranges.Where(rg => rangeClauseDO.MinValue.IsWithin(rg.Item1, rg.Item2));
            var overlapsMax = summaryCoverage.Ranges.Where(rg => rangeClauseDO.MaxValue.IsWithin(rg.Item1, rg.Item2));
            foreach (var rg in summaryCoverage.Ranges)
            {
                if (overlapsMin.Contains(rg))
                {
                    updatedRanges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(rg.Item1, rangeClauseDO.MaxValue));
                }
                else if (overlapsMax.Contains(rg))
                {
                    updatedRanges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(rangeClauseDO.MinValue, rg.Item2));
                }
                else
                {
                    updatedRanges.Add(rg);
                }
            }

            if (!overlapsMin.Any() && !overlapsMax.Any())
            {
                updatedRanges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(rangeClauseDO.MinValue, rangeClauseDO.MaxValue));
            }

            summaryCoverage.Ranges = updatedRanges;

            summaryCoverage = AggregateSummaryRanges(summaryCoverage);

            return summaryCoverage;
        }

        private SummaryCaseCoverage UpdateSummaryDataSingleValues(SummaryCaseCoverage summaryClauses, RangeClauseDataObject rangeClauseDO)
        {
            if (!rangeClauseDO.CanBeInspected || rangeClauseDO.SingleValue == null) { return summaryClauses; }

            if (rangeClauseDO.UsesIsClause)
            {
                if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.LT)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.LTE)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GT)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GTE))
                {
                    summaryClauses = UpdateSummaryIsClauseLimits(rangeClauseDO.SingleValue, rangeClauseDO.CompareSymbol, summaryClauses);
                }
                else if (CompareTokens.EQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    summaryClauses.SingleValues.Add(rangeClauseDO.SingleValue);
                }
                else if (CompareTokens.NEQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    summaryClauses = UpdateSummaryIsClauseLimits(rangeClauseDO.SingleValue, CompareTokens.LT, summaryClauses);
                    summaryClauses = UpdateSummaryIsClauseLimits(rangeClauseDO.SingleValue, CompareTokens.GT, summaryClauses);
                }
            }
            else
            {
                summaryClauses.SingleValues.Add(rangeClauseDO.SingleValue);
            }
            return summaryClauses;
        }

        private bool EvaluateCaseElseAccessibility(SummaryCaseCoverage summaryClauses, string typeName)
        {
            if (summaryClauses.CaseElseIsUnreachable) { return summaryClauses.CaseElseIsUnreachable; }

            if (typeName.Equals(Tokens.Boolean))
            {
                return summaryClauses.SingleValues.Any(val => val == UnreachableCaseInspectionValue.Zero) && summaryClauses.SingleValues.Any(val => val != UnreachableCaseInspectionValue.Zero)
                    || summaryClauses.IsLT != null && summaryClauses.IsLT > UnreachableCaseInspectionValue.False
                    || summaryClauses.IsGT != null && summaryClauses.IsGT < UnreachableCaseInspectionValue.True
                    || summaryClauses.IsLT != null && summaryClauses.IsLT == UnreachableCaseInspectionValue.False && summaryClauses.SingleValues.Any(sv => sv == UnreachableCaseInspectionValue.False)
                    || summaryClauses.IsGT != null && summaryClauses.IsGT == UnreachableCaseInspectionValue.True && summaryClauses.SingleValues.Any(sv => sv == UnreachableCaseInspectionValue.True);
            }

            if (typeName.Equals(Tokens.Byte))
            {
                if (summaryClauses.SingleValues.Count >= UnreachableCaseInspectionValue.MaxValueByte + 1)
                {
                    summaryClauses.SingleValues = summaryClauses.SingleValues.Distinct().ToList();
                    return summaryClauses.SingleValues.Count() == UnreachableCaseInspectionValue.MaxValueByte + 1;
                }
            }

            if (summaryClauses.IsLT != null && summaryClauses.IsGT != null)
            {
                if (summaryClauses.IsLT > summaryClauses.IsGT
                        || (summaryClauses.IsLT >= summaryClauses.IsGT
                                && summaryClauses.SingleValues.Contains(summaryClauses.IsLT)))
                {
                    return true;
                }

                else if (summaryClauses.Ranges.Count > 0)
                {
                    if (!IsIntegerNumberType(summaryClauses.IsLT.UseageTypeName))
                    {
                        return false;
                    }

                    var remainingValues = new List<long>();
                    for (var idx = summaryClauses.IsLT.AsLong().Value; idx <= summaryClauses.IsGT.AsLong().Value; idx++)
                    {
                        remainingValues.Add(idx);
                    }
                    remainingValues.RemoveAll(rv => summaryClauses.Ranges.Any(rg => rg.Item1.AsLong().Value <= rv && rg.Item2.AsLong().Value >= rv));
                    if (remainingValues.Any())
                    {
                        remainingValues.RemoveAll(rv => summaryClauses.SingleValues.Contains(new UnreachableCaseInspectionValue(rv, Tokens.Long)));
                        return !remainingValues.Any();
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private SummaryCaseCoverage AggregateSummaryRanges(SummaryCaseCoverage currentSummaryCaseCoverage)
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

        private List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> AppendRanges(List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> ranges)
        {
            if (ranges.Count() <= 1 || !IsIntegerNumberType(ranges.First().Item1.UseageTypeName))
            {
                return ranges;
            }

            var updatedRanges = new List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>();
            var combinedLastRange = false;

            for (var idx = 0; idx < ranges.Count(); idx++)
            {
                if (idx + 1 >= ranges.Count())
                {
                    if (!combinedLastRange)
                    {
                        updatedRanges.Add(ranges[idx]);
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
                    updatedRanges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(theMin, theNextMax));
                    combinedLastRange = true;
                }
                else if (theMin.AsLong() == theNextMax.AsLong() + 1)
                {
                    updatedRanges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(theNextMin, theMax));
                    combinedLastRange = true;
                }
                else
                {
                    updatedRanges.Add(ranges[idx]);
                }
            }
            return updatedRanges;
        }

        private static ExpressionEvaluationDataObject GetEvaluationData(IParseTree ctxt, Dictionary<IParseTree, ExpressionEvaluationDataObject> ctxtEvalResults)
        {
            return ctxtEvalResults.ContainsKey(ctxt) ? ctxtEvalResults[ctxt] : new ExpressionEvaluationDataObject { Operator = string.Empty, CanBeInspected = true };
        }

        private static Dictionary<IParseTree, ExpressionEvaluationDataObject> AddEvaluationData(Dictionary<IParseTree, ExpressionEvaluationDataObject> contextIndices, IParseTree ctxt, ExpressionEvaluationDataObject exprEvaluation)
        {
            if (contextIndices.ContainsKey(ctxt))
            {
                contextIndices[ctxt] = exprEvaluation;
            }
            else
            {
                contextIndices.Add(ctxt, exprEvaluation);
            }
            return contextIndices;
        }

        private static bool IsIntegerNumberType(string typeName) => new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte }.Contains(typeName);

        private static bool TryGetChildContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : ParserRuleContext
        {
            opCtxt = ctxt.FindChild<U>();
            return opCtxt != null;
        }

        private static List<UnreachableCaseInspectionValue> LoadRangeOfByteValues(List<UnreachableCaseInspectionValue> SingleValues, long start, long end)
        {
            if (start >= UnreachableCaseInspectionValue.MinValueByte 
                    && start <= UnreachableCaseInspectionValue.MaxValueByte
                    && start <= end)
            {
                var constrainedEnd = end >= UnreachableCaseInspectionValue.MaxValueByte ? UnreachableCaseInspectionValue.MaxValueByte : end;
                for (var val = start; val <= constrainedEnd; val++)
                {
                    SingleValues.Add(new UnreachableCaseInspectionValue(val));
                }
            }
            return SingleValues;
        }

        private bool IsBinaryOperatorContext<T>(T child)
        {
            return IsBinaryMathOperation(child)
                || IsBinaryLogicalOperation(child);
        }

        private bool IsMathOperation<T>(T child)
        {
            return IsBinaryMathOperation(child)
                || IsUnaryMathOperation(child);
        }

        private bool IsBinaryMathOperation<T>(T child)
        {
            return child is MultOpContext
                || child is AddOpContext
                || child is PowOpContext
                || child is ModOpContext;
        }

        private bool IsBinaryLogicalOperation<T>(T child)
        {
            return child is RelationalOpContext
                || child is LogicalXorOpContext
                || child is LogicalAndOpContext
                || child is LogicalOrOpContext
                || child is LogicalEqvOpContext
                || child is LogicalNotOpContext;
        }

        private bool IsUnaryLogicalOperator<T>(T child)
        {
            return child is LogicalNotOpContext;
        }

        private bool IsUnaryMathOperation<T>(T child)
        {
            return child is UnaryMinusOpContext;
        }

        private bool IsUnaryOperandContext<T>(T child)
        {
            return IsUnaryLogicalOperator(child)
                   || IsUnaryMathOperation(child)
                   || child is ParenthesizedExprContext;
        }

        private bool ContextCanBeEvaluated(ParserRuleContext context, string refName)
        {
            var canBeInspected = true;
            var descs = context.GetDescendents();
            var ops = descs.Where(desc => (desc is ParserRuleContext) && (IsBinaryMathOperation(desc) || IsUnaryMathOperation(desc)));
            foreach (var op in ops)
            {
                var lExpressions = ((ParserRuleContext)op).FindChildren<LExprContext>();
                var mathOnTheSelectCaseVariable = lExpressions.Any(lex => lex.GetText().Equals(refName));
                var mathOnNonConstants = lExpressions.Any(lex => !(CreateValue(lex, Tokens.Variant).HasValue));

                if (mathOnTheSelectCaseVariable || mathOnNonConstants)
                {                
                    canBeInspected = false;
                }
            }
            return canBeInspected;
        }

        private RangeClauseDataObject SetTheCompareOperator(RangeClauseDataObject rangeClauseDO)
        {
            rangeClauseDO.UsesIsClause = TryGetChildContext(rangeClauseDO.Context, out ComparisonOperatorContext opCtxt);
            rangeClauseDO.CompareSymbol = rangeClauseDO.UsesIsClause ? opCtxt.GetText() : CompareTokens.EQ;
            return rangeClauseDO;
        }

        private SelectStmtDataObject SetTheTypeNames(SelectStmtDataObject selectStmtDO, string typeName, string baseTypeName = "")
        {
            selectStmtDO.AsTypeName = typeName;
            selectStmtDO.BaseTypeName = baseTypeName.Length == 0 ? typeName : baseTypeName;
            return selectStmtDO;
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

            public override void EnterSelectCaseStmt([NotNull] SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
        #endregion
    }
}