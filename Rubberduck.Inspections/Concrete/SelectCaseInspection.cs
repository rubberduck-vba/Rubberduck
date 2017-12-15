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
    //TODO: Add replace with UI Resource
    public static class CaseInspectionMessages
    {
        public static string Unreachable => "Unreachable Case Statement";
        public static string MismatchType => "Type does not match the Select Statement";
        public static string CaseElse => "All possible values are handled by prior Case statement(s)";
    }

    public sealed class SelectCaseInspection : ParseTreeInspectionBase
    {

        internal enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse, NoResult };

        internal Dictionary<ClauseEvaluationResult, string> _resultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            { ClauseEvaluationResult.Unreachable, CaseInspectionMessages.Unreachable },
            { ClauseEvaluationResult.MismatchType, CaseInspectionMessages.MismatchType },
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

        private static Dictionary<string, Func<VBEValue, VBEValue, VBEValue>> ValueOperations = new Dictionary<string, Func<VBEValue, VBEValue, VBEValue>>()
        {
            { "*", delegate(VBEValue LHS, VBEValue RHS){ return LHS * RHS; } },
            { "/", delegate(VBEValue LHS, VBEValue RHS){ return LHS / RHS; } },
            { "+", delegate(VBEValue LHS, VBEValue RHS){ return LHS + RHS; } },
            { "-", delegate(VBEValue LHS, VBEValue RHS){ return LHS - RHS; } },
            { "^", delegate(VBEValue LHS, VBEValue RHS){ return LHS ^ RHS; } }
        };

        private static Dictionary<string, Func<VBEValue, VBEValue, VBEValue>> CompareOperations = new Dictionary<string, Func<VBEValue, VBEValue, VBEValue>>()
        {
            { CompareSymbols.EQ, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue(LHS == RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareSymbols.NEQ, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue(LHS != RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareSymbols.LT, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue(LHS < RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareSymbols.LTE, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue(LHS <= RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareSymbols.GT, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue(LHS > RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareSymbols.GTE, delegate(VBEValue LHS, VBEValue RHS){ return new VBEValue(LHS >= RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } }
        };

        internal struct SummaryCaseCoverage
        {
            public VBEValue IsLTMax;
            public VBEValue IsGTMin;
            public List<VBEValue> SingleValues;
            public List<Tuple<VBEValue, VBEValue>> Ranges;
            public List<string> Indeterminants;
            public List<string> RangeClausesAsText;
            public Dictionary<string, List<Tuple<long, bool>>> Discretes;
        }

        internal struct ExpressionEval
        {
            public ParserRuleContext ParentCtxt;
            public bool SingleValueOnly;
            public ExpressionContext LHS;
            public Object LHSEval;
            public VBEValue LHSValue;
            public ExpressionContext RHS;
            public Object RHSEval;
            public VBEValue RHSValue;
            public string Operator;
            public VBEValue Result;
            public bool CanBeInspected;
        }

        internal struct SelectStmtDataObject
        {
            public QualifiedContext<ParserRuleContext> QualifiedCtxt;
            public string BaseTypeName;
            public string AsTypeName;
            public IdentifierReference IdReference;
            public List<CaseClauseDataObject> CaseClauseDOs;
            public CaseElseClauseContext CaseElseContext;
            public SummaryCaseCoverage SummaryClauses;
            public bool HasUnreachableCaseElse;
            public bool CanBeInspected;
            public Dictionary<string, long> EnumerationValues;

            public SelectStmtDataObject(QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                QualifiedCtxt = selectStmtCtxt;
                BaseTypeName = string.Empty;
                AsTypeName = string.Empty;
                IdReference = null;
                CaseClauseDOs = new List<CaseClauseDataObject>();
                CaseElseContext = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(QualifiedCtxt.Context);
                HasUnreachableCaseElse = false;
                CanBeInspected = true;
                EnumerationValues = null;
                SummaryClauses = new SummaryCaseCoverage
                {
                    IsGTMin = null,
                    IsLTMax = null,
                    SingleValues = new List<VBEValue>(),
                    Ranges = new List<Tuple<VBEValue, VBEValue>>(),
                    Indeterminants = new List<string>(),
                    RangeClausesAsText = new List<string>(),
                    Discretes = new Dictionary<string, List<Tuple<long, bool>>>()
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
            public RangeClauseContext Context;
            public bool UsesIsClause;
            public bool IsValueRange;
            public bool IsConstant;
            public bool IsParseable;
            public bool CompareByTextOnly;
            public string IdReferenceName;
            public string AsText;
            public string TypeNameNative;
            public string TypeNameTarget;
            public string CompareSymbol;
            public VBEValue SingleValue;
            public VBEValue MinValue;
            public VBEValue MaxValue;
            public ClauseEvaluationResult ResultType;
            public bool CanBeInspected;

            public RangeClauseDataObject(RangeClauseContext ctxt)
            {
                Context = ctxt;
                UsesIsClause = false;
                IsValueRange = false;
                IsConstant = false;
                IsParseable = false;
                CompareByTextOnly = false;
                IdReferenceName = string.Empty;
                AsText = ctxt.GetText();
                TypeNameNative = string.Empty;
                TypeNameTarget = string.Empty;
                CompareSymbol = CompareSymbols.EQ;
                SingleValue = null;
                MinValue = null;
                MaxValue = null;
                ResultType = ClauseEvaluationResult.NoResult;
                CanBeInspected = true;
            }
        }

        public SelectCaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        private static Dictionary<string, string> CompareInversions = new Dictionary<string, string>()
        {
            { CompareSymbols.EQ, CompareSymbols.EQ },
            { CompareSymbols.NEQ, CompareSymbols.NEQ },
            { CompareSymbols.LT, CompareSymbols.GT },
            { CompareSymbols.LTE, CompareSymbols.GTE },
            { CompareSymbols.GT, CompareSymbols.LT },
            { CompareSymbols.GTE, CompareSymbols.LTE }
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var inspResults = new List<IInspectionResult>();

            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            foreach (var selectStmt in selectCaseContexts)
            {
                var selectStmtDO = InitializeSelectStatementDataObject(new SelectStmtDataObject(selectStmt));
                if (!selectStmtDO.BaseTypeName.Equals(string.Empty) && selectStmtDO.CanBeInspected)
                {
                    for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count; idx++)
                    {
                        var caseClauseDO = selectStmtDO.CaseClauseDOs[idx];
                        if(caseClauseDO.ResultType != ClauseEvaluationResult.Unreachable)
                        {
                            for (var rgIdx = 0; rgIdx < caseClauseDO.RangeClauseDOs.Count; rgIdx++)
                            {
                                var rgClause = caseClauseDO.RangeClauseDOs[rgIdx];
                                rgClause = InitializeRangeClauseDataObject(rgClause, selectStmtDO.BaseTypeName, selectStmtDO.IdReference);
                                caseClauseDO.RangeClauseDOs[rgIdx] = rgClause;
                            }
                        }
                        caseClauseDO.ResultType = caseClauseDO.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.MismatchType)
                            ? ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;
                        selectStmtDO.CaseClauseDOs[idx] = caseClauseDO;
                    }
                    selectStmtDO = InspectSelectStmtCaseClauses(selectStmtDO);

                    inspResults.AddRange(selectStmtDO.CaseClauseDOs.Where(cc => cc.ResultType != ClauseEvaluationResult.NoResult)
                        .Select(cc => CreateInspectionResult(selectStmt, cc.CaseContext, _resultMessages[cc.ResultType])));

                    if (selectStmtDO.HasUnreachableCaseElse && selectStmtDO.CaseElseContext != null)
                    {
                        inspResults.Add(CreateInspectionResult(selectStmt, selectStmtDO.CaseElseContext, _resultMessages[ClauseEvaluationResult.CaseElse]));
                    }
                }
            }
            return inspResults;
        }

        private SelectStmtDataObject InitializeSelectStatementDataObject(SelectStmtDataObject selectStmtDO)
        {
            if (TryGetChildContext(selectStmtDO.QualifiedCtxt.Context, out SelectExpressionContext selectExprCtxt))
            {
                selectStmtDO = ResolveContextType(selectStmtDO, selectExprCtxt);
            }
            else
            {
                selectStmtDO.CanBeInspected = false;
                return selectStmtDO;
            }

            if (selectStmtDO.BaseTypeName.Equals(Tokens.Variant))
            {
                if (TryInferVariantTypeName(selectStmtDO, out string typeName))
                {
                    selectStmtDO.BaseTypeName = typeName;
                    selectStmtDO.AsTypeName = typeName;
                }
                else
                {
                    selectStmtDO.CanBeInspected = false;
                }
            }

            if (!selectStmtDO.BaseTypeName.Equals(string.Empty) && selectStmtDO.CanBeInspected)
            {
                var caseClauseCtxts = ParserRuleContextHelper.GetChildren<CaseClauseContext>(selectStmtDO.QualifiedCtxt.Context);
                selectStmtDO.CaseClauseDOs = caseClauseCtxts.Select(cc => CreateCaseClauseDataObject(cc)).ToList();
            }
            return selectStmtDO;
        }

        private bool TryInferVariantTypeName(SelectStmtDataObject selectStmtDO, out string typeName)
        {
            typeName = selectStmtDO.BaseTypeName;
            if (selectStmtDO.BaseTypeName.Equals(Tokens.Variant))
            {
                var caseClauseCtxts = ParserRuleContextHelper.GetChildren<CaseClauseContext>(selectStmtDO.QualifiedCtxt.Context);
                var rangeCtxts = caseClauseCtxts.Select(cc => ParserRuleContextHelper.GetChildren<RangeClauseContext>(cc)).SelectMany(rg => rg);
                var typeNames = new List<string>();
                foreach (var rg in rangeCtxts)
                {
                    typeNames.AddRange(ParserRuleContextHelper.GetDescendents<LiteralExprContext>(rg).Select(de => EvaluateRangeClauseTypeName(GetText(de), Tokens.Variant)));
                }
                if (typeNames.All(tn => typeNames.First().Equals(tn)))
                {
                    typeName = typeNames.First();
                    return true;
                }
                if (typeNames.Contains(Tokens.String))
                {
                    //typeName = Tokens.String;
                    return false;
                }
                if (typeNames.Contains(Tokens.Double))
                {
                    typeName = Tokens.Double;
                    return true;
                }
                if (typeNames.Contains(Tokens.Long))
                {
                    typeName = Tokens.Long;
                    return true;
                }
            }
            return false;
        }

        private CaseClauseDataObject CreateCaseClauseDataObject(CaseClauseContext ctxt)
        {
            var caseClauseDO = new CaseClauseDataObject(ctxt);
            var rangeClauseContexts = ParserRuleContextHelper.GetChildren<RangeClauseContext>(ctxt);
            foreach (var rangeClauseCtxt in rangeClauseContexts)
            {
                caseClauseDO.RangeClauseDOs.Add(new RangeClauseDataObject(rangeClauseCtxt));
            }
            return caseClauseDO;
        }

        private RangeClauseDataObject InitializeRangeClauseDataObject(RangeClauseDataObject rangeClauseDO, string targetTypeName, IdentifierReference idRef)
        {
            rangeClauseDO.TypeNameTarget = targetTypeName;
            rangeClauseDO.TypeNameNative = targetTypeName;
            rangeClauseDO.IdReferenceName = idRef != null ? idRef.IdentifierName : string.Empty;
            rangeClauseDO.UsesIsClause = HasChildToken(rangeClauseDO.Context, Tokens.Is);
            rangeClauseDO.IsValueRange = HasChildToken(rangeClauseDO.Context, Tokens.To);
            rangeClauseDO = SetTheCompareOperator(rangeClauseDO);

            if (rangeClauseDO.IsValueRange)
            {
                var startContext = ParserRuleContextHelper.GetChild<SelectStartValueContext>(rangeClauseDO.Context);
                var endContext = ParserRuleContextHelper.GetChild<SelectEndValueContext>(rangeClauseDO.Context);
                var startValue = ResolveContextValue(ref rangeClauseDO, startContext);
                var endValue = ResolveContextValue(ref rangeClauseDO, endContext);

                var startTypeName = EvaluateRangeClauseTypeName(startContext.GetText(), rangeClauseDO.TypeNameTarget);
                var endTypeName = EvaluateRangeClauseTypeName(endContext.GetText(), rangeClauseDO.TypeNameTarget);

                if (!startTypeName.Equals(endTypeName))
                {
                    //Find common ground for comparisons if possible
                    if (startTypeName.Equals(Tokens.String) || endTypeName.Equals(Tokens.String))
                    {
                        //Forcing comparisons as strings is not reliable for numbers
                        rangeClauseDO.TypeNameNative = string.Empty;
                        if (!(startValue.IsParseable && endValue.IsParseable))
                        {
                            rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                            return rangeClauseDO;
                        }
                    }
                    else if (startTypeName.Equals(Tokens.Double) || endTypeName.Equals(Tokens.Double))
                    {
                        startValue = new VBEValue(startValue.AsString(), Tokens.Double);
                        endValue = new VBEValue(endValue.AsString(), Tokens.Double);
                    }
                    else if (startTypeName.Equals(Tokens.Long) || endTypeName.Equals(Tokens.Long))
                    {
                        startValue = new VBEValue(startValue.AsString(), Tokens.Long);
                        endValue = new VBEValue(endValue.AsString(), Tokens.Long);
                    }
                    else
                    {
                        rangeClauseDO.TypeNameNative = string.Empty;
                        if (!(startValue.IsParseable && endValue.IsParseable))
                        {
                            rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                            return rangeClauseDO;
                        }
                    }
                }
                rangeClauseDO.MinValue = startValue <= endValue ? startValue : endValue;
                rangeClauseDO.MaxValue = startValue <= endValue ? endValue : startValue;
                rangeClauseDO.SingleValue = rangeClauseDO.MinValue;
                rangeClauseDO.IsParseable = rangeClauseDO.MinValue.IsParseable && rangeClauseDO.MaxValue.IsParseable;
            }
            else
            {
                rangeClauseDO.TypeNameNative = EvaluateRangeClauseTypeName(rangeClauseDO.Context.GetText(), rangeClauseDO.TypeNameTarget);
                rangeClauseDO.SingleValue = ResolveContextValue(ref rangeClauseDO, rangeClauseDO.Context);
                rangeClauseDO.MaxValue = rangeClauseDO.MinValue;
                rangeClauseDO.MinValue = rangeClauseDO.SingleValue;
                rangeClauseDO.IsParseable = rangeClauseDO.SingleValue == null ? false : rangeClauseDO.SingleValue.IsParseable;
            }

            rangeClauseDO.CompareByTextOnly = !rangeClauseDO.IsParseable && rangeClauseDO.TypeNameNative.Equals(rangeClauseDO.TypeNameTarget);
            rangeClauseDO.ResultType = !rangeClauseDO.TypeNameTarget.Equals(rangeClauseDO.TypeNameNative) && !rangeClauseDO.IsParseable ?
                    ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;

            return rangeClauseDO;
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

        private RangeClauseDataObject SetTheCompareOperator(RangeClauseDataObject rangeClauseDO)
        {
            rangeClauseDO.UsesIsClause = TryGetChildContext(rangeClauseDO.Context, out ComparisonOperatorContext opCtxt);
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

        private string EvaluateRangeClauseTypeName(string textValue, string targetTypeName)
        {
            //TODO use TypeHintToTypeName - and add tests for each kind
            if (SymbolList.TypeHintToTypeName.TryGetValue(textValue.Last().ToString(), out string typeName))
            {
                return typeName;
            }

            if (IsStringLiteral(textValue))
            {
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
                return targetTypeName;
            }
            else if (textValue.Equals(Tokens.True) || textValue.Equals(Tokens.False))
            {
                return Tokens.Boolean;
            }
            else if(long.TryParse(textValue, out _))
            {
                return Tokens.Long;
            }
            else
            {
                return targetTypeName;
            }
        }
        
        private static string GetText(ParserRuleContext ctxt) => ctxt.GetText().Replace("\"", "");

        private VBEValue ResolveContextValue(ref RangeClauseDataObject rangeClauseDO, ParserRuleContext context)
        {
            var eval = ResolveContextValue(ref rangeClauseDO, new Dictionary<ParserRuleContext, ExpressionEval>(), context);
            rangeClauseDO.CompareSymbol = eval[context].Operator;
            return eval[context].Result;
        }

        private Dictionary<ParserRuleContext,ExpressionEval> ResolveContextValue(ref RangeClauseDataObject rangeClauseDO, Dictionary<ParserRuleContext, ExpressionEval> contextEvals, ParserRuleContext parentContext)
        {
            if(parentContext is RangeClauseContext 
                || parentContext is SelectStartValueContext 
                || parentContext is SelectEndValueContext )
            {
                var parentEval = new ExpressionEval { SingleValueOnly = true };
                contextEvals = AddEvalData(contextEvals, parentContext, parentEval);
            }

            foreach(var child in parentContext.children)
            {
                if (!rangeClauseDO.CanBeInspected || child is WhiteSpaceContext) { continue; }

                if ( child is RelationalOpContext
                    || child is MultOpContext
                    || child is AddOpContext
                    || child is ParenthesizedExprContext
                    || child is PowOpContext
                    || child is LogicalXorOpContext
                    || child is LogicalAndOpContext
                    || child is LogicalOrOpContext
                    || child is LogicalEqvOpContext
                    || child is LogicalNotOpContext
                    )
                {
                    var childData = GetEvalData((ParserRuleContext)child, contextEvals);
                    childData.ParentCtxt = parentContext;
                    childData.SingleValueOnly = child is ParenthesizedExprContext;

                    var lExprCtxts = ParserRuleContextHelper.GetChildren<LExprContext>((RuleContext)child);
                    var refName = rangeClauseDO.IdReferenceName;
                   
                    if (lExprCtxts.Any(lex => lex.GetText().Equals(refName)))
                    {
                        //TODO: implement the necessary algebra to support these use cases
                        if (child is MultOpContext
                            || child is AddOpContext
                            || child is PowOpContext
                            )
                        {
                            rangeClauseDO.CanBeInspected = false;
                            childData.CanBeInspected = false;
                        }
                    }

                    rangeClauseDO.UsesIsClause = child is RelationalOpContext
                        || child is LogicalXorOpContext
                        || child is LogicalAndOpContext
                        || child is LogicalOrOpContext
                        || child is LogicalEqvOpContext
                        || child is LogicalNotOpContext;

                    contextEvals = AddEvalData(contextEvals, (ParserRuleContext)child, childData);
                    if (rangeClauseDO.CanBeInspected)
                    {
                        contextEvals = ResolveContextValue(ref rangeClauseDO, contextEvals, (ParserRuleContext)child);
                    }
                    contextEvals = UpdateParentEvaluation(ref rangeClauseDO, (ParserRuleContext)child, contextEvals);
                }
                else if (child is LiteralExprContext 
                    || child is LExprContext 
                    || child is UnaryMinusOpContext)
                {
                    var childData = GetEvalData((ParserRuleContext)child, contextEvals);
                    childData.ParentCtxt = parentContext;
                    childData.SingleValueOnly = true;
                    childData.LHS = (ExpressionContext)child;
                    childData.LHSValue = EvaluateContextValue(childData.LHS, rangeClauseDO.TypeNameTarget);
                    childData.Result = childData.LHSValue;

                    contextEvals = AddEvalData(contextEvals, (ParserRuleContext)child, childData);
                    contextEvals = UpdateParentEvaluation(ref rangeClauseDO, (ParserRuleContext)child, contextEvals);
                }
                else if (ValueOperations.Keys.Contains(child.GetText()) || CompareOperations.Keys.Contains(child.GetText()))
                {
                    var parentData = GetEvalData(parentContext, contextEvals);
                    parentData.Operator = child.GetText();
                    contextEvals = AddEvalData(contextEvals, parentContext, parentData);
                }
            }
            return contextEvals;
        }

        private Dictionary<ParserRuleContext, ExpressionEval> UpdateParentEvaluation(ref RangeClauseDataObject rgClauseDO, ParserRuleContext child, Dictionary<ParserRuleContext, ExpressionEval> ctxtEvalResults)
        {
            var childData = ctxtEvalResults[child];
            var parentContext = (ParserRuleContext)child.Parent;
            if (!childData.CanBeInspected)
            {
                var parentData = GetEvalData(parentContext, ctxtEvalResults);
                parentData.CanBeInspected = false;
                ctxtEvalResults = AddEvalData(ctxtEvalResults, parentContext, parentData);
                return ctxtEvalResults;
            }
            if (childData.Result != null)
            {
                var parentData = GetEvalData(parentContext, ctxtEvalResults);

                if(childData.Operator != null && CompareOperations.ContainsKey(childData.Operator))
                {
                    parentData.Operator = childData.Operator;
                }

                if (parentData.SingleValueOnly)
                {
                    if (parentData.LHS == null)
                    {
                        parentData.LHS = (ExpressionContext)child;
                        parentData.LHSValue = childData.Result;
                        parentData.Result = parentData.LHSValue;
                    }
                }
                else
                {
                    if (parentData.LHS == null)
                    {
                        parentData.LHS = (ExpressionContext)child;
                        parentData.LHSValue = childData.Result;
                    }
                    else
                    {
                        parentData.RHS = (ExpressionContext)child;
                        parentData.RHSValue = childData.Result;
                    }

                    if (parentData.LHSValue != null && parentData.RHSValue != null)
                    {
                        if(parentData.Operator == string.Empty || parentData.Operator == null)
                        {
                            parentData.Result = null;
                        }
                        else
                        {
                            var LHS = parentData.LHSValue;
                            var RHS = parentData.RHSValue;
                            if (RHS.AsString().Equals(rgClauseDO.IdReferenceName))
                            {
                                LHS = RHS;
                                RHS = parentData.LHSValue;
                                parentData.Operator = CompareInversions[parentData.Operator];
                            }
                            var result = GetOpExpressionResult(LHS, RHS, parentData.Operator);
                            parentData.Result = new VBEValue(result, rgClauseDO.TypeNameTarget);
                        }
                    }
                }
                ctxtEvalResults = AddEvalData(ctxtEvalResults, parentContext, parentData);
            }
            return ctxtEvalResults;
        }

        private ExpressionEval GetEvalData(ParserRuleContext ctxt, Dictionary<ParserRuleContext, ExpressionEval> contextIndices)
        {
            return contextIndices.ContainsKey(ctxt) ? contextIndices[ctxt] : new ExpressionEval { Operator = string.Empty, CanBeInspected = true };
        }

        private Dictionary<ParserRuleContext, ExpressionEval> AddEvalData(Dictionary<ParserRuleContext, ExpressionEval> contextIndices, ParserRuleContext ctxt, ExpressionEval idx)
        {
            if (contextIndices.ContainsKey(ctxt))
            {
                contextIndices[ctxt] = idx;
            }
            else
            {
                contextIndices.Add(ctxt, idx);
            }
            return contextIndices;
        }

        private VBEValue EvaluateContextValue(ExpressionContext ctxt, string typeName)
        {
            if (ctxt is LExprContext)
            {
                if(TryGetTheExpressionValue((LExprContext)ctxt, out string lexpr))
                {
                    return new VBEValue(lexpr, typeName);
                }
                return new VBEValue(ctxt.GetText(), typeName);
            }
            else if (ctxt is LiteralExprContext)
            {
                return new VBEValue(GetText((LiteralExprContext)ctxt), typeName);
            }
            else if (ctxt is UnaryMinusOpContext)
            {
                return new VBEValue(GetText((UnaryMinusOpContext)ctxt), typeName);
            }
            return null;
        }

        private string GetOpExpressionResult(VBEValue LHS, VBEValue RHS, string operation)
        {
            if (ValueOperations.ContainsKey(operation))
            {
                return ValueOperations[operation](LHS, RHS).AsString();
            }
            else if (CompareOperations.ContainsKey(operation))
            {
                //The parseable check is needed to support cases like 'x < 4' - where 'x' 
                //is the SelectCase variable
                //TODO: a better way?
                return LHS.IsParseable ? CompareOperations[operation](LHS, RHS).AsString() : RHS.AsString();
            }
            return string.Empty;
        }

        private bool TryGetTheExpressionValue(LExprContext ctxt, out string expressionValue)
        {
            expressionValue = string.Empty;
            var member = ParserRuleContextHelper.GetChild<MemberAccessExprContext>(ctxt);
            if (member != null)
            {
                var smplNameMemberLHS = ParserRuleContextHelper.GetChild<SimpleNameExprContext>(member);
                var smplNameMemberRHS = ParserRuleContextHelper.GetChild<UnrestrictedIdentifierContext>(member);
                var theLHSName = smplNameMemberLHS.GetText();
                var theRHSName = smplNameMemberRHS.GetText();
                var memberDeclarations = State.DeclarationFinder.AllUserDeclarations.Where(dec => dec.IdentifierName.Equals(theRHSName));

                foreach (var dec in memberDeclarations)
                {
                    if (dec.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        var theCtxt = dec.Context;
                        if (theCtxt is EnumerationStmt_ConstantContext)
                        {
//////                            var contextEvals = ResolveContextValue(ref rangeClauseDO, new Dictionary<ParserRuleContext, ExpressionEval>(), theCtxt);
                            var valuedDeclaration = ParserRuleContextHelper.GetChild<LiteralExprContext>(theCtxt);
                            expressionValue = valuedDeclaration.GetText();
                            return true;
                        }
                    }
                }
                return false;
            }

            var smplName = ParserRuleContextHelper.GetChild<SimpleNameExprContext>(ctxt);
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

        private IdentifierReference GetTheRangeClauseReference(SimpleNameExprContext smplNameCtxt, string theName)
        {
            var identifierReferences = (State.DeclarationFinder.MatchName(theName).Select(dec => dec.References)).SelectMany(rf => rf);

            if (!identifierReferences.Any())
            {
                return null;
            }

            if (identifierReferences.Count() == 1)
            {
                return identifierReferences.First();
            }
            else
            {
                var rangeClauseReference = identifierReferences.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, smplNameCtxt)
                                        && (ParserRuleContextHelper.HasParent(rf.Context, smplNameCtxt.Parent)));

                Debug.Assert(rangeClauseReference.Count() == 1);
                return rangeClauseReference.First();
            }
        }

        private SelectStmtDataObject ResolveContextType(SelectStmtDataObject selectStmtDO, ParserRuleContext parentCtxt)
        {
            if (TryGetChildContext(parentCtxt, out RelationalOpContext _)
                || TryGetChildContext(parentCtxt, out LogicalXorOpContext _)
                || TryGetChildContext(parentCtxt, out LogicalAndOpContext _)
                || TryGetChildContext(parentCtxt, out LogicalOrOpContext _)
                || TryGetChildContext(parentCtxt, out LogicalEqvOpContext _)
                || TryGetChildContext(parentCtxt, out LogicalNotOpContext _)
                )
            {
                selectStmtDO.AsTypeName = Tokens.Boolean;
                selectStmtDO.BaseTypeName = Tokens.Boolean;
            }
            else if (TryGetChildContext(parentCtxt, out MultOpContext mCtxt))
            {
                return ResolveContextType(selectStmtDO, mCtxt);
            }
            else if (TryGetChildContext(parentCtxt, out AddOpContext aCtxt))
            {
                return ResolveContextType(selectStmtDO, aCtxt);
            }
            else if (TryGetChildContext(parentCtxt, out PowOpContext pCtxt))
            {
                return ResolveContextType(selectStmtDO, pCtxt);
            }
            else if (TryGetChildContext(parentCtxt, out ParenthesizedExprContext parenCtxt))
            {
                return ResolveContextType(selectStmtDO, parenCtxt);
            }
            else if (TryGetChildContext(parentCtxt, out UnaryMinusOpContext unaryCtxt))
            {
                return ResolveContextType(selectStmtDO, unaryCtxt);
            }
            else if (TryGetChildContext(parentCtxt, out LExprContext lexpr))
            {
                var idCs = ParserRuleContextHelper.GetDescendents<IdentifierContext>(lexpr);
                var idRefs = idCs.Select(idc => GetTheSelectCaseReference(parentCtxt.Parent, idCs.First().GetText()));
                if (idRefs.Any())
                {
                    selectStmtDO.IdReference = idRefs.First();
                    selectStmtDO.AsTypeName = selectStmtDO.IdReference.Declaration.AsTypeName;
                    var isConsistentType = idRefs.All(idr => idr.Declaration.AsTypeName == selectStmtDO.AsTypeName);
                    if (selectStmtDO.IdReference.Declaration.AsTypeIsBaseType)
                    {
                        selectStmtDO.BaseTypeName = selectStmtDO.AsTypeName;
                    }
                    else if (selectStmtDO.IdReference.Declaration.AsTypeDeclaration.AsTypeIsBaseType)
                    {
                        selectStmtDO.BaseTypeName = selectStmtDO.IdReference.Declaration.AsTypeDeclaration.AsTypeName;
                    }
                }
            }
            return selectStmtDO;
        }

        #region EnumerationDeadCode
        //else
        //{
        //    var smplNames = ParserRuleContextHelper.GetDescendents<SimpleNameExprContext>(selectExprCtxt);
        //    if (smplNames.Any())
        //    {
        //        foreach (var smplName in smplNames)
        //        {
        //            selectStmtDO.IdReference = GetTheSelectCaseReference(selectStmtDO.QualifiedCtxt.Context, smplName.GetText());
        //            if (selectStmtDO.IdReference != null)
        //            {
        //                selectStmtDO.AsTypeName = selectStmtDO.IdReference.Declaration.AsTypeName;
        //                if (selectStmtDO.IdReference.Declaration.AsTypeIsBaseType)
        //                {
        //                    selectStmtDO.BaseTypeName = selectStmtDO.IdReference.Declaration.AsTypeName;
        //                }
        //                else if (selectStmtDO.IdReference.Declaration.AsTypeDeclaration.AsTypeIsBaseType)
        //                {
        //                    selectStmtDO.BaseTypeName = selectStmtDO.IdReference.Declaration.AsTypeDeclaration.AsTypeName;
        //                    //TODO: how to ensure the declaration is the right one?
        //                    var decID = State.DeclarationFinder.AllUserDeclarations.Where(dec => dec.IdentifierName.Equals(selectStmtDO.AsTypeName));
        //                    var theDec = decID.Any() ? decID.First() : null;  //hack
        //                    if (theDec != null)
        //                    {
        //                        if (theDec.DeclarationType.HasFlag(DeclarationType.Enumeration))
        //                        {
        //                            selectStmtDO.EnumerationValues = new Dictionary<string, long>();
        //                            var enumMembers = ParserRuleContextHelper.GetChildren<EnumerationStmt_ConstantContext>(theDec.Context);
        //                            foreach (var enumMember in enumMembers)
        //                            {
        //                                var theIDCtxt = ParserRuleContextHelper.GetChild<IdentifierContext>(enumMember);
        //                                var theValueCtxt = ParserRuleContextHelper.GetChild<LiteralExprContext>(enumMember);
        //                                if (theValueCtxt != null)
        //                                {
        //                                    var expressionValue = theValueCtxt.GetText();
        //                                    var enumID = theIDCtxt.GetText();
        //                                    //TODO: the long.Parse may fail if the declaration is something like B = A * 2
        //                                    if (selectStmtDO.EnumerationValues != null)
        //                                    {
        //                                        selectStmtDO.EnumerationValues.Add(enumID, long.Parse(expressionValue));
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    selectStmtDO.EnumerationValues = null;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //                return selectStmtDO;
        //            }
        //        }
        //    }
        //}
        //return selectStmtDO;
#endregion

        private IdentifierReference GetTheSelectCaseReference(RuleContext selectCaseStmtCtxt, string theName)
        {
            var identifierReferences = (State.DeclarationFinder.MatchName(theName).Select(dec => dec.References)).SelectMany(rf => rf);

            //TODO: Is there a scenario that results in two or more results?
            return identifierReferences.Any() ? identifierReferences.First() : null;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private SelectStmtDataObject InspectSelectStmtCaseClauses(SelectStmtDataObject selectStmtDO)
        {
            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = selectStmtDO.CaseClauseDOs[idx];
                if (selectStmtDO.HasUnreachableCaseElse)
                {
                    //If the CaseElse is previously determined to be unreachable, 
                    //then all of the yet-to-be-evaluated CaseClauses unreachable as well.
                    caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                }
                else
                {
                    selectStmtDO = InspectCaseClause(ref caseClause, selectStmtDO);
                    selectStmtDO.HasUnreachableCaseElse = caseClause.MakesRemainingClausesUnreachable;
                    if(caseClause.ResultType == ClauseEvaluationResult.NoResult)
                    {
                        if (caseClause.RangeClauseDOs.All(rg => selectStmtDO.SummaryClauses.RangeClausesAsText.Contains(rg.AsText)))
                        {
                            caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                        }
                    }
                    selectStmtDO.SummaryClauses.RangeClausesAsText.AddRange(caseClause.RangeClauseDOs.Select(rg => rg.AsText));
                }
                selectStmtDO.CaseClauseDOs[idx] = caseClause;
            }

            return selectStmtDO;
        }

        private static bool ExceedsMinMax(RangeClauseDataObject rangeClauseDO)
        {
            if (rangeClauseDO.IsValueRange)
            {
                return rangeClauseDO.MinValue.ExceedsMaxMin() && rangeClauseDO.MaxValue.ExceedsMaxMin();
            }
            return rangeClauseDO.SingleValue.ExceedsMaxMin();
        }

        private SelectStmtDataObject InspectRangeClause(SelectStmtDataObject selectStmtDO, ref RangeClauseDataObject rangeClauseDO)
        {
            if(rangeClauseDO.ResultType != ClauseEvaluationResult.NoResult)
            {
                return selectStmtDO;
            }

            if (ExceedsMinMax(rangeClauseDO))
            {
                rangeClauseDO.ResultType = ClauseEvaluationResult.Unreachable;
                return selectStmtDO;
            }

            if (!rangeClauseDO.IsValueRange)
            {
                if (rangeClauseDO.UsesIsClause)
                {
                    if (new string[] { CompareSymbols.LT, CompareSymbols.LTE, CompareSymbols.GT, CompareSymbols.GTE }.Contains(rangeClauseDO.CompareSymbol))
                    {
                        selectStmtDO.SummaryClauses = UpdateSummaryIsClauseLimits(rangeClauseDO.SingleValue, rangeClauseDO.CompareSymbol, selectStmtDO.SummaryClauses);
                    }
                    else if (CompareSymbols.EQ.Equals(rangeClauseDO.CompareSymbol))
                    {
                        selectStmtDO.SummaryClauses = HandleSimpleSingleValueCompare(ref rangeClauseDO, rangeClauseDO.SingleValue, selectStmtDO.SummaryClauses);
                    }
                    else if (CompareSymbols.NEQ.Equals(rangeClauseDO.CompareSymbol))
                    {
                        if (selectStmtDO.SummaryClauses.SingleValues.Contains(rangeClauseDO.SingleValue))
                        {
                            rangeClauseDO.ResultType = ClauseEvaluationResult.CaseElse;
                        }

                        selectStmtDO.SummaryClauses = UpdateSummaryIsClauseLimits(rangeClauseDO.SingleValue, CompareSymbols.LT, selectStmtDO.SummaryClauses);
                        selectStmtDO.SummaryClauses = UpdateSummaryIsClauseLimits(rangeClauseDO.SingleValue, CompareSymbols.GT, selectStmtDO.SummaryClauses);
                    }
                }
                else
                {
                    selectStmtDO.SummaryClauses = HandleSimpleSingleValueCompare(ref rangeClauseDO, rangeClauseDO.SingleValue, selectStmtDO.SummaryClauses);
                }
            }
            else  //It is a range of values like "Case 45 To 150"
            {
                selectStmtDO.SummaryClauses = AggregateRanges(selectStmtDO.SummaryClauses);
                var minValue = rangeClauseDO.MinValue;
                var maxValue = rangeClauseDO.MaxValue;
                rangeClauseDO.ResultType = selectStmtDO.SummaryClauses.Ranges.Where(rg => minValue.IsWithin(rg.Item1, rg.Item2)
                        && maxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                        || selectStmtDO.SummaryClauses.IsLTMax != null && selectStmtDO.SummaryClauses.IsLTMax > rangeClauseDO.MaxValue
                        || selectStmtDO.SummaryClauses.IsGTMin != null && selectStmtDO.SummaryClauses.IsGTMin < rangeClauseDO.MinValue
                        ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

                if (rangeClauseDO.ResultType == ClauseEvaluationResult.NoResult)
                {
                    var overlapsMin = selectStmtDO.SummaryClauses.Ranges.Where(rg => minValue.IsWithin(rg.Item1, rg.Item2));
                    var overlapsMax = selectStmtDO.SummaryClauses.Ranges.Where(rg => maxValue.IsWithin(rg.Item1, rg.Item2));
                    var updated = new List<Tuple<VBEValue, VBEValue>>();
                    foreach (var rg in selectStmtDO.SummaryClauses.Ranges)
                    {
                        if (overlapsMin.Contains(rg))
                        {
                            updated.Add(new Tuple<VBEValue, VBEValue>(rg.Item1, rangeClauseDO.MaxValue));
                        }
                        else if (overlapsMax.Contains(rg))
                        {
                            updated.Add(new Tuple<VBEValue, VBEValue>(rangeClauseDO.MinValue, rg.Item2));
                        }
                        else
                        {
                            updated.Add(rg);
                        }
                    }

                    if (!overlapsMin.Any() && !overlapsMax.Any())
                    {
                        updated.Add(new Tuple<VBEValue, VBEValue>(rangeClauseDO.MinValue, rangeClauseDO.MaxValue));
                    }
                    selectStmtDO.SummaryClauses.Ranges = updated;
                }

                if (selectStmtDO.BaseTypeName.Equals(Tokens.Boolean))
                {
                    rangeClauseDO.ResultType = rangeClauseDO.MinValue != rangeClauseDO.MaxValue
                        ? ClauseEvaluationResult.CaseElse : ClauseEvaluationResult.NoResult;
                }
            }
            return selectStmtDO;
        }

        private SelectStmtDataObject InspectCaseClause(ref CaseClauseDataObject caseClause, SelectStmtDataObject selectStmtDO)
        {
            if(caseClause.ResultType != ClauseEvaluationResult.NoResult)
            {
                return selectStmtDO;
            }

            for (var idx = 0; idx < caseClause.RangeClauseDOs.Count(); idx++)
            {
                var range = caseClause.RangeClauseDOs[idx];
                if (range.CanBeInspected)
                {
                    selectStmtDO = InspectRangeClause(selectStmtDO, ref range);
                }
                caseClause.RangeClauseDOs[idx] = range;
            }

            caseClause.MakesRemainingClausesUnreachable =
                    caseClause.RangeClauseDOs.Where(rg => rg.ResultType == ClauseEvaluationResult.CaseElse).Any()
                    || IsClausesCoverAllValues(selectStmtDO.SummaryClauses);

            caseClause.ResultType = caseClause.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.Unreachable)
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            return selectStmtDO;
        }

        //TODO: change this to evaluate the SummaryClauses for any and all types of full coverage
        private bool IsClausesCoverAllValues(SummaryCaseCoverage summaryClauses)
        {
            if (summaryClauses.IsLTMax != null && summaryClauses.IsGTMin != null)
            {
                return summaryClauses.IsLTMax > summaryClauses.IsGTMin
                        || (summaryClauses.IsLTMax >= summaryClauses.IsGTMin
                        && summaryClauses.SingleValues.Contains(summaryClauses.IsLTMax));
            }
            return false;
        }

        private bool SingleValueIsHandledPreviously(VBEValue theValue, SummaryCaseCoverage priorHandlers)
        {
            return priorHandlers.IsLTMax != null && theValue < priorHandlers.IsLTMax
                || priorHandlers.IsGTMin != null && theValue > priorHandlers.IsGTMin
                || priorHandlers.SingleValues.Contains(theValue)
                || priorHandlers.Ranges.Where(rg => theValue.IsWithin(rg.Item1, rg.Item2)).Any();
        }

        private SummaryCaseCoverage UpdateSummaryIsClauseLimits(VBEValue theValue, string compareSymbol, SummaryCaseCoverage priorHandlers)
        {
            if (compareSymbol.Equals(CompareSymbols.LT) || compareSymbol.Equals(CompareSymbols.LTE) ) // new string[] { CompareSymbols.LT, CompareSymbols.LTE }.Contains(compareSymbol))
            {
                priorHandlers.IsLTMax = priorHandlers.IsLTMax == null ? theValue
                    : priorHandlers.IsLTMax < theValue ? theValue : priorHandlers.IsLTMax;
            }
            else if (compareSymbol.Equals(CompareSymbols.GT) || compareSymbol.Equals(CompareSymbols.GTE))
            {
                priorHandlers.IsGTMin = priorHandlers.IsGTMin == null ? theValue
                    : priorHandlers.IsGTMin > theValue ? theValue : priorHandlers.IsGTMin;
            }
            else
            {
                return priorHandlers;
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

        private SummaryCaseCoverage HandleSimpleSingleValueCompare(ref RangeClauseDataObject range, VBEValue theValue, SummaryCaseCoverage priorHandlers)
        {
            range.ResultType = SingleValueIsHandledPreviously(theValue, priorHandlers)
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            if (theValue.TargetTypeName.Equals(Tokens.Boolean))
            {
                range.ResultType = priorHandlers.SingleValues.Any()
                    && !priorHandlers.SingleValues.Contains(theValue)
                    ? ClauseEvaluationResult.CaseElse : range.ResultType;
            }

            if (range.ResultType != ClauseEvaluationResult.Unreachable)
            {
                priorHandlers.SingleValues.Add(theValue);
            }
            return priorHandlers;
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

        private List<Tuple<VBEValue, VBEValue>> AppendRanges(List<Tuple<VBEValue, VBEValue>> ranges)
        {
            if (ranges.Count() <= 1)
            {
                return ranges;
            }

            if (!ranges.First().Item1.IsIntegerNumber)
            {
                return ranges;
            }

            var updatedHandlers = new List<Tuple<VBEValue, VBEValue>>();
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
                    updatedHandlers.Add(new Tuple<VBEValue, VBEValue>(theMin, theNextMax));
                    combinedLastRange = true;
                }
                else if (theMin.AsLong() == theNextMax.AsLong() + 1)
                {
                    updatedHandlers.Add(new Tuple<VBEValue, VBEValue>(theNextMin, theMax));
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