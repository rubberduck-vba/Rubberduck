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

        internal static class CompareTokens
        {
            public static readonly string EQ = "=";
            public static readonly string NEQ = "<>";
            public static readonly string LT = "<";
            public static readonly string LTE = "<=";
            public static readonly string GT = ">";
            public static readonly string GTE = ">=";
        }

        private static Dictionary<string, Func<VBAValue, VBAValue, VBAValue>> MathOperations = new Dictionary<string, Func<VBAValue, VBAValue, VBAValue>>()
        {
            { "*", delegate(VBAValue LHS, VBAValue RHS){ return LHS * RHS; } },
            { "/", delegate(VBAValue LHS, VBAValue RHS){ return LHS / RHS; } },
            { "+", delegate(VBAValue LHS, VBAValue RHS){ return LHS + RHS; } },
            { "-", delegate(VBAValue LHS, VBAValue RHS){ return LHS - RHS; } },
            { "^", delegate(VBAValue LHS, VBAValue RHS){ return LHS ^ RHS; } },
            { "Mod", delegate(VBAValue LHS, VBAValue RHS){ return LHS % RHS; } }
        };

        private static Dictionary<string, Func<VBAValue, VBAValue, VBAValue>> CompareOperations = new Dictionary<string, Func<VBAValue, VBAValue, VBAValue>>()
        {
            { CompareTokens.EQ, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue(LHS == RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareTokens.NEQ, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue(LHS != RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareTokens.LT, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue(LHS < RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareTokens.LTE, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue(LHS <= RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareTokens.GT, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue(LHS > RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } },
            { CompareTokens.GTE, delegate(VBAValue LHS, VBAValue RHS){ return new VBAValue(LHS >= RHS ? Tokens.True: Tokens.False, Tokens.Boolean); } }
        };

        internal struct SummaryCaseCoverage
        {
            public VBAValue IsLTMax;
            public VBAValue IsGTMin;
            public List<VBAValue> SingleValues;
            public List<Tuple<VBAValue, VBAValue>> Ranges;
            public bool UnreachableCaseElse;
            public List<string> RangeClausesAsText;
            public List<VBAValue> EnumDiscretes;
        }

        internal struct ExpressionEval
        {
            public ParserRuleContext ParentCtxt;
            public bool IsUnaryOperation;
            public VBAValue LHSValue;
            public VBAValue RHSValue;
            public string Operator;
            public string SelectCaseRefName;
            public string TypeNameTarget;
            public VBAValue Result;
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
            //public Dictionary<string, long> EnumerationValues;

            public SelectStmtDataObject(QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                SelectStmtContext = (SelectCaseStmtContext)selectStmtCtxt.Context;
                IdReferenceName = string.Empty;
                BaseTypeName = Tokens.Variant;
                AsTypeName = Tokens.Variant;
                CaseClauseDOs = new List<CaseClauseDataObject>();
                CaseElseContext = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(SelectStmtContext);
                //EnumerationValues = null;
                SummaryCaseClauses = new SummaryCaseCoverage
                {
                    IsGTMin = null,
                    IsLTMax = null,
                    SingleValues = new List<VBAValue>(),
                    Ranges = new List<Tuple<VBAValue, VBAValue>>(),
                    RangeClausesAsText = new List<string>(),
                    EnumDiscretes = new List<VBAValue>()
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
            public bool IsParseable;
            public bool CompareByTextOnly;
            public string IdReferenceName;
            public string AsText;
            public string TypeNameDerived;
            public string TypeNameTarget;
            public string CompareSymbol;
            public VBAValue SingleValue;
            public VBAValue MinValue;
            public VBAValue MaxValue;
            public ClauseEvaluationResult ResultType;
            public bool CanBeInspected;

            public RangeClauseDataObject(RangeClauseContext ctxt, string targetTypeName)
            {
                Context = ctxt;
                UsesIsClause = false;
                IsValueRange = false;
                IsConstant = false;
                IsParseable = false;
                CompareByTextOnly = false;
                IdReferenceName = string.Empty;
                AsText = ctxt.GetText();
                TypeNameDerived = string.Empty;
                TypeNameTarget = targetTypeName;
                CompareSymbol = CompareTokens.EQ;
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
            { CompareTokens.EQ, CompareTokens.EQ },
            { CompareTokens.NEQ, CompareTokens.NEQ },
            { CompareTokens.LT, CompareTokens.GT },
            { CompareTokens.LTE, CompareTokens.GTE },
            { CompareTokens.GT, CompareTokens.LT },
            { CompareTokens.GTE, CompareTokens.LTE }
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var inspResults = new List<IInspectionResult>();

            var selectCaseContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            foreach (var selectStmt in selectCaseContexts)
            {
                var selectStmtDO = InitializeSelectStatementDataObject(new SelectStmtDataObject(selectStmt));

                if ( !selectStmtDO.CanBeInspected) { continue; }

                selectStmtDO = InspectSelectStmtCaseClauses(selectStmtDO);

                inspResults.AddRange(selectStmtDO.CaseClauseDOs.Where(cc => cc.ResultType != ClauseEvaluationResult.NoResult)
                    .Select(cc => CreateInspectionResult(selectStmt, cc.CaseContext, _resultMessages[cc.ResultType])));

                if (selectStmtDO.SummaryCaseClauses.UnreachableCaseElse && selectStmtDO.CaseElseContext != null)
                {
                    inspResults.Add(CreateInspectionResult(selectStmt, selectStmtDO.CaseElseContext, _resultMessages[ClauseEvaluationResult.CaseElse]));
                }
            }
            return inspResults;
        }

        private SelectStmtDataObject InitializeSelectStatementDataObject(SelectStmtDataObject selectStmtDO)
        {
            selectStmtDO = ResolveSelectStmtType(selectStmtDO);

            if (selectStmtDO.BaseTypeName.Equals(Tokens.Variant))
            {
                if (TryInferAnalysisTypeForVariant( selectStmtDO, out string typeName))
                {
                    selectStmtDO = SetAsAndBaseType(selectStmtDO, typeName);
                }
                else
                {
                    selectStmtDO.CanBeInspected = false;
                    return selectStmtDO;
                }
            }

            if (selectStmtDO.CanBeInspected)
            {
                selectStmtDO.CaseClauseDOs = ParserRuleContextHelper.GetChildren<CaseClauseContext>(selectStmtDO.SelectStmtContext)
                    .Select(cc => CreateCaseClauseDataObject(cc, selectStmtDO.BaseTypeName)).ToList();
            }
            return selectStmtDO;
        }

        private bool TryInferAnalysisTypeForVariant(SelectStmtDataObject selectStmtDO, out string typeName)
        {
            typeName = selectStmtDO.BaseTypeName;
            if (selectStmtDO.BaseTypeName.Equals(Tokens.Variant))
            {
                var rangeCtxts = ParserRuleContextHelper.GetChildren<CaseClauseContext>(selectStmtDO.SelectStmtContext)
                    .Select(cc => ParserRuleContextHelper.GetChildren<RangeClauseContext>(cc))
                    .SelectMany(rg => rg);

                var typeNames = rangeCtxts.SelectMany(rg => ParserRuleContextHelper.GetDescendents(rg))
                    .Where(desc => desc is LiteralExprContext || desc is LExprContext)
                        .Select(exprCtxt => EvaluateContextTypeName((ExpressionContext)exprCtxt, selectStmtDO.SelectStmtContext))
                        .Where(tp => tp != string.Empty);

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
            }
            return false;
        }

        private CaseClauseDataObject CreateCaseClauseDataObject(CaseClauseContext ctxt, string targetTypeName)
        {
            var caseClauseDO = new CaseClauseDataObject(ctxt);
            var rangeClauseContexts = ParserRuleContextHelper.GetChildren<RangeClauseContext>(ctxt);
            foreach (var rangeClauseCtxt in rangeClauseContexts)
            {
                var rgC = new RangeClauseDataObject(rangeClauseCtxt, targetTypeName);
                caseClauseDO.RangeClauseDOs.Add(rgC);
            }
            return caseClauseDO;
        }

        private RangeClauseDataObject InitializeRangeClauseDataObject(RangeClauseDataObject rangeClauseDO, string targetTypeName, string refName )
        {
            rangeClauseDO.TypeNameTarget = targetTypeName;
            rangeClauseDO.TypeNameDerived = targetTypeName;
            rangeClauseDO.IdReferenceName = refName;
            rangeClauseDO.UsesIsClause = HasChildToken(rangeClauseDO.Context, Tokens.Is);
            rangeClauseDO.IsValueRange = HasChildToken(rangeClauseDO.Context, Tokens.To);
            rangeClauseDO = SetTheCompareOperator(rangeClauseDO);

            if (rangeClauseDO.IsValueRange)
            {
                //rangeClauseDO =  InitializeValueRangeClauseDataObject(rangeClauseDO, targetTypeName, refName);
                var startContext = ParserRuleContextHelper.GetChild<SelectStartValueContext>(rangeClauseDO.Context);
                var endContext = ParserRuleContextHelper.GetChild<SelectEndValueContext>(rangeClauseDO.Context);
                var startEnd = new Tuple<VBAValue, VBAValue>(ResolveRangeClauseValue(ref rangeClauseDO, startContext), ResolveRangeClauseValue(ref rangeClauseDO, endContext));

                var startTypeName = startEnd.Item1.HasValue ? startEnd.Item1.UseageTypeName : Tokens.String;
                var endTypeName = startEnd.Item2.HasValue ? startEnd.Item2.UseageTypeName : Tokens.String;

                if (!startTypeName.Equals(endTypeName))
                {
                    var typePrecedence = new string[] { Tokens.Double, Tokens.Long, Tokens.Integer, Tokens.Byte };
                    //Find common ground for comparisons if possible
                    if (startTypeName.Equals(Tokens.String) || endTypeName.Equals(Tokens.String))
                    {
                        //Forcing comparisons as strings is not reliable for numbers
                        rangeClauseDO.TypeNameDerived = string.Empty;
                        if (!(startEnd.Item1.HasValue && startEnd.Item2.HasValue))
                        {
                            rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                            return rangeClauseDO;
                        }
                    }
                    else if (typePrecedence.Contains(startTypeName) || typePrecedence.Contains(endTypeName))
                    {
                        for (var idx = 0; idx < typePrecedence.Count(); idx++)
                        {
                            if (typePrecedence[idx].Equals(startTypeName) || typePrecedence[idx].Equals(endTypeName))
                            {
                                var newStart = new VBAValue(startEnd.Item1.AsString(), typePrecedence[idx]);
                                var newEnd = new VBAValue(startEnd.Item2.AsString(), typePrecedence[idx]);
                                startEnd = new Tuple<VBAValue, VBAValue>(newStart, newEnd);
                                rangeClauseDO.TypeNameDerived = typePrecedence[idx];
                                idx = typePrecedence.Count();
                            }
                        }
                    }
                    else
                    {
                        rangeClauseDO.TypeNameDerived = string.Empty;
                        if (!(startEnd.Item1.HasValue && startEnd.Item2.HasValue))
                        {
                            rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                            return rangeClauseDO;
                        }
                    }
                }
                rangeClauseDO.MinValue = startEnd.Item1 <= startEnd.Item2 ? startEnd.Item1 : startEnd.Item2;
                rangeClauseDO.MaxValue = startEnd.Item1 <= startEnd.Item2 ? startEnd.Item2 : startEnd.Item1;
                rangeClauseDO.SingleValue = rangeClauseDO.MinValue;
                rangeClauseDO.IsParseable = rangeClauseDO.MinValue.HasValue && rangeClauseDO.MaxValue.HasValue;
            }
            else
            {
                rangeClauseDO.TypeNameDerived = VBAValue.DeriveTypeName(rangeClauseDO.Context.GetText(), rangeClauseDO.TypeNameTarget);
                if(IsIntegerNumberType(rangeClauseDO.TypeNameDerived) && IsIntegerNumberType(rangeClauseDO.TypeNameTarget))
                {
                    rangeClauseDO.TypeNameDerived = rangeClauseDO.TypeNameTarget;
                }
                rangeClauseDO.SingleValue = ResolveRangeClauseValue(ref rangeClauseDO, rangeClauseDO.Context);
                rangeClauseDO.MaxValue = rangeClauseDO.MinValue;
                rangeClauseDO.MinValue = rangeClauseDO.SingleValue;
                rangeClauseDO.IsParseable = rangeClauseDO.SingleValue == null ? false : rangeClauseDO.SingleValue.HasValue;
            }

            rangeClauseDO.CompareByTextOnly = !rangeClauseDO.IsParseable && rangeClauseDO.TypeNameDerived.Equals(rangeClauseDO.TypeNameTarget);
            rangeClauseDO.ResultType = !rangeClauseDO.TypeNameTarget.Equals(rangeClauseDO.TypeNameDerived) && !rangeClauseDO.IsParseable ?
                    ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;

            return rangeClauseDO;
        }

        private RangeClauseDataObject InitializeValueRangeClauseDataObject(RangeClauseDataObject rangeClauseDO, string targetTypeName, string refName)
        {
            var startContext = ParserRuleContextHelper.GetChild<SelectStartValueContext>(rangeClauseDO.Context);
            var endContext = ParserRuleContextHelper.GetChild<SelectEndValueContext>(rangeClauseDO.Context);
            var startEnd = new Tuple<VBAValue, VBAValue>(ResolveRangeClauseValue(ref rangeClauseDO, startContext), ResolveRangeClauseValue(ref rangeClauseDO, endContext));

            var startTypeName = startEnd.Item1.HasValue ? startEnd.Item1.UseageTypeName : Tokens.String;
            var endTypeName = startEnd.Item2.HasValue ? startEnd.Item2.UseageTypeName : Tokens.String;

            if (!startTypeName.Equals(endTypeName))
            {
                var typePrecedence = new string[] { Tokens.Double, Tokens.Long, Tokens.Integer, Tokens.Byte };
                //Find common ground for comparisons if possible
                if (startTypeName.Equals(Tokens.String) || endTypeName.Equals(Tokens.String))
                {
                    //Forcing comparisons as strings is not reliable for numbers
                    rangeClauseDO.TypeNameDerived = string.Empty;
                    if (!(startEnd.Item1.HasValue && startEnd.Item2.HasValue))
                    {
                        rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                        return rangeClauseDO;
                    }
                }
                else if (typePrecedence.Contains(startTypeName) || typePrecedence.Contains(endTypeName))
                {
                    for (var idx = 0; idx < typePrecedence.Count(); idx++)
                    {
                        if (typePrecedence[idx].Equals(startTypeName) || typePrecedence[idx].Equals(endTypeName))
                        {
                            var newStart = new VBAValue(startEnd.Item1.AsString(), typePrecedence[idx]);
                            var newEnd = new VBAValue(startEnd.Item2.AsString(), typePrecedence[idx]);
                            startEnd = new Tuple<VBAValue, VBAValue>(newStart, newEnd);
                            rangeClauseDO.TypeNameDerived = typePrecedence[idx];
                            idx = typePrecedence.Count();
                        }
                    }
                }
                else
                {
                    rangeClauseDO.TypeNameDerived = string.Empty;
                    if (!(startEnd.Item1.HasValue && startEnd.Item2.HasValue))
                    {
                        rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                        return rangeClauseDO;
                    }
                }
            }
            rangeClauseDO.MinValue = startEnd.Item1 <= startEnd.Item2 ? startEnd.Item1 : startEnd.Item2;
            rangeClauseDO.MaxValue = startEnd.Item1 <= startEnd.Item2 ? startEnd.Item2 : startEnd.Item1;
            rangeClauseDO.SingleValue = rangeClauseDO.MinValue;
            rangeClauseDO.IsParseable = rangeClauseDO.MinValue.HasValue && rangeClauseDO.MaxValue.HasValue;

            return rangeClauseDO;
        }
        private VBAValue ResolveRangeClauseValue(ref RangeClauseDataObject rangeClauseDO, ParserRuleContext context)
        {
            if (!(context is RangeClauseContext || context is SelectStartValueContext || context is SelectEndValueContext))
            {
                return null;
            }

            var parentEval = new ExpressionEval
            {
                IsUnaryOperation = true,
                Operator = rangeClauseDO.CompareSymbol,
                CanBeInspected = rangeClauseDO.CanBeInspected,
                TypeNameTarget = rangeClauseDO.TypeNameTarget,
                SelectCaseRefName = rangeClauseDO.IdReferenceName
            };

            var contextEvals = AddEvaluationData(new Dictionary<ParserRuleContext, ExpressionEval>(), context, parentEval);

            contextEvals = ResolveContextValue(contextEvals, context);
            rangeClauseDO.CompareSymbol = contextEvals[context].Operator;
            rangeClauseDO.UsesIsClause = rangeClauseDO.UsesIsClause ? rangeClauseDO.UsesIsClause : contextEvals[context].EvaluateAsIsClause;
            return contextEvals[context].Result;
        }

        private Dictionary<ParserRuleContext,ExpressionEval> ResolveContextValue( Dictionary<ParserRuleContext, ExpressionEval> contextEvals, ParserRuleContext parentContext)
        {
            foreach (var child in parentContext.children)
            {
                if (child is WhiteSpaceContext) { continue; }

                var parentData = GetEvaluationData(parentContext, contextEvals);

                if (MathOperations.Keys.Contains(child.GetText()) || CompareOperations.Keys.Contains(child.GetText()))
                {
                    if (!parentData.EvaluateAsIsClause)
                    {
                        parentData.EvaluateAsIsClause = CompareOperations.Keys.Contains(child.GetText());
                    }
                    parentData.Operator = child.GetText();
                    contextEvals = AddEvaluationData(contextEvals, parentContext, parentData);
                    continue;
                }

                if (IsBinaryOperatorContext(child) || IsUnaryOperandContext( child) )
                {
                    var childData = GetEvaluationData((ParserRuleContext)child, contextEvals);
                    childData.ParentCtxt = parentContext;
                    childData.IsUnaryOperation = IsUnaryOperandContext(child);
                    childData.TypeNameTarget = parentData.TypeNameTarget;
                    childData.SelectCaseRefName = parentData.SelectCaseRefName;

                    if (!childData.EvaluateAsIsClause)
                    {
                        childData.EvaluateAsIsClause = IsBinaryLogicalOperation(child) || IsUnaryLogicalOperator(child);
                    }

                    contextEvals = AddEvaluationData(contextEvals, (ParserRuleContext)child, childData);
                    contextEvals = ResolveContextValue(contextEvals, (ParserRuleContext)child);
                    contextEvals = UpdateParentEvaluation((ParserRuleContext)child, contextEvals);
                }
                else if (child is LiteralExprContext || child is LExprContext )
                {
                    var childData = GetEvaluationData((ParserRuleContext)child, contextEvals);
                    childData.ParentCtxt = parentContext;
                    childData.IsUnaryOperation = true;
                    childData.TypeNameTarget = parentData.TypeNameTarget;
                    childData.SelectCaseRefName = parentData.SelectCaseRefName;
                    childData.LHSValue = EvaluateContextValue((ExpressionContext)child, childData.TypeNameTarget); // rangeClauseDO.TypeNameTarget);
                    childData.Result = childData.LHSValue;

                    contextEvals = AddEvaluationData(contextEvals, (ParserRuleContext)child, childData);
                    contextEvals = UpdateParentEvaluation((ParserRuleContext)child, contextEvals);
                }
            }
            return contextEvals;
        }

        private Dictionary<ParserRuleContext, ExpressionEval> UpdateParentEvaluation(ParserRuleContext child, Dictionary<ParserRuleContext, ExpressionEval> ctxtEvalResults)
        {
            var childData = ctxtEvalResults[child];
            var parentData = GetEvaluationData(childData.ParentCtxt, ctxtEvalResults);

            if (!childData.CanBeInspected)
            {
                parentData.CanBeInspected = false;
                ctxtEvalResults = AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
                return ctxtEvalResults;
            }

            if (!parentData.EvaluateAsIsClause)
            {
                parentData.EvaluateAsIsClause  = childData.EvaluateAsIsClause;
            }

            if (childData.Result == null)
            {
                return ctxtEvalResults;
            }

            if(childData.Operator != null && CompareOperations.ContainsKey(childData.Operator))
            {
                parentData.Operator = childData.Operator;
            }

            if (parentData.IsUnaryOperation)
            {
                parentData.LHSValue = childData.Result;
                if (childData.ParentCtxt is UnaryMinusOpContext )
                {
                    var inverseOperand = new VBAValue(-1, parentData.LHSValue.UseageTypeName);
                    parentData.LHSValue = parentData.LHSValue * inverseOperand;
                }
                parentData.Result = parentData.LHSValue;
            }
            else
            {
                if (parentData.LHSValue == null)
                {
                    parentData.LHSValue = childData.Result;
                }
                else if(parentData.Operator != string.Empty || parentData.Operator != null)
                {
                    parentData.RHSValue = childData.Result;

                    //For cases like '45 > x', flip around (invert) the operation to
                    //'x < 45' so it conforms to 'Is' statement format ( 'Is < 45' ) and can be treated as such
                    var invertOperation = parentData.RHSValue.AsString().Equals(childData.SelectCaseRefName);

                    parentData.Operator = invertOperation ? CompareInversions[parentData.Operator] : parentData.Operator;
                    var exprResult = invertOperation ?
                            GetOpExpressionResult(parentData.RHSValue, parentData.LHSValue, parentData.Operator)
                            : GetOpExpressionResult(parentData.LHSValue, parentData.RHSValue, parentData.Operator);

                    parentData.Result = new VBAValue(exprResult, childData.TypeNameTarget);
                }
            }
            return AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
        }

        private ExpressionEval GetEvaluationData(ParserRuleContext ctxt, Dictionary<ParserRuleContext, ExpressionEval> ctxtEvalResults)
        {
            return ctxtEvalResults.ContainsKey(ctxt) ? ctxtEvalResults[ctxt] : new ExpressionEval { Operator = string.Empty, CanBeInspected = true };
        }

        private Dictionary<ParserRuleContext, ExpressionEval> AddEvaluationData(Dictionary<ParserRuleContext, ExpressionEval> contextIndices, ParserRuleContext ctxt, ExpressionEval exprEvaluation)
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

        private string EvaluateContextTypeName(ExpressionContext ctxt, SelectCaseStmtContext selectStmtContext)
        {
            if (ctxt is LiteralExprContext)
            {
                return VBAValue.DeriveTypeName(GetText(ctxt));
            }
            else if (ctxt is LExprContext)
            {
                return ResolveContextType(selectStmtContext, ctxt as LExprContext);
            }
            return string.Empty;
        }

        private VBAValue EvaluateContextValue(ExpressionContext ctxt, string typeName)
        {
            if (ctxt is LExprContext)
            {
                if(TryGetTheExpressionValue((LExprContext)ctxt, out string lexpr))
                {
                    return new VBAValue(lexpr, typeName);
                }
                //TODO: Should this return null? - this is what executes when it is a non-constant reference
                return new VBAValue(ctxt.GetText(), typeName);
            }
            else if (ctxt is LiteralExprContext)
            {
                return new VBAValue(GetText((LiteralExprContext)ctxt), typeName);
            }
            return null;
        }

        private string GetOpExpressionResult(VBAValue LHS, VBAValue RHS, string operation)
        {
            if (MathOperations.ContainsKey(operation))
            {
                return MathOperations[operation](LHS, RHS).AsString();
            }
            else if (CompareOperations.ContainsKey(operation))
            {
                //Supports cases like 'x < 4' - where 'x' 
                //is the SelectCase variable
                //TODO: a better way?
                return LHS.HasValue ? CompareOperations[operation](LHS, RHS).AsString() : RHS.AsString();
            }
            return string.Empty;
        }

        private bool TryGetTheExpressionValue(LExprContext ctxt, out string expressionValue)
        {
            expressionValue = string.Empty;
            var member = ParserRuleContextHelper.GetChild<MemberAccessExprContext>(ctxt);
            if (member != null)
            {
                var smplNameMemberRHS = ParserRuleContextHelper.GetChild<UnrestrictedIdentifierContext>(member);
                var memberDeclarations = State.DeclarationFinder.AllUserDeclarations.Where(dec => dec.IdentifierName.Equals(smplNameMemberRHS.GetText()));

                foreach (var dec in memberDeclarations)
                {
                    if (dec.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        var theCtxt = dec.Context;
                        if (theCtxt is EnumerationStmt_ConstantContext)
                        {
                            expressionValue = GetConstantDeclarationValue(dec);
                            return true;
                        }
                    }
                }
                return false;
            }

            var smplName = ParserRuleContextHelper.GetChild<SimpleNameExprContext>(ctxt);
            if (smplName != null)
            {
                var identifierReferences = (State.DeclarationFinder.MatchName(smplName.GetText()).Select(dec => dec.References)).SelectMany(rf => rf);
                var rangeClauseReferences = identifierReferences.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, smplName)
                                        && (ParserRuleContextHelper.HasParent(rf.Context, smplName.Parent)));

                var rangeClauseIdentifierReference = rangeClauseReferences.Any() ? rangeClauseReferences.First() : null;
                if (rangeClauseIdentifierReference != null)
                {
                    if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant)
                        || rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        expressionValue = GetConstantDeclarationValue(rangeClauseIdentifierReference.Declaration);
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
                    var parentData = new Dictionary<ParserRuleContext, ExpressionEval>();
                    var exprEval = new ExpressionEval
                    {
                        IsUnaryOperation = IsUnaryMathOperation(child),
                        Operator = "=",
                        CanBeInspected = true,
                        TypeNameTarget = valueDeclaration.AsTypeName,
                        SelectCaseRefName = valueDeclaration.IdentifierName
                    };

                    parentData = AddEvaluationData(parentData, child, exprEval);
                    return ResolveContextValue(parentData, child).First().Value.Result.AsString();
                }
                if (child is LiteralExprContext)
                {
                    if(child.Parent is EnumerationStmt_ConstantContext)
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
            var foundEqualSign = false;
            for (int idx = 0; idx < contexts.Count(); idx++)
            {
                var childCtxt = contexts[idx];
                if (childCtxt.GetText().Equals("="))
                {
                    foundEqualSign = true;
                    continue;
                }
                if (foundEqualSign && childCtxt is ParserRuleContext && !(childCtxt is WhiteSpaceContext))
                {
                    contextsOfInterest.Add((ParserRuleContext)childCtxt);
                }
            }
            return contextsOfInterest;
        }

        private SelectStmtDataObject ResolveSelectStmtType(SelectStmtDataObject selectStmtDO )//, ParserRuleContext ctxt)
        {
            var parentCtxt = selectStmtDO.SelectExpressionContext;

            if (parentCtxt.ChildCount == 1 && IsBinaryLogicalOperation(parentCtxt.children[0]))
            {
                return SetAsAndBaseType(selectStmtDO, Tokens.Boolean);
            }

            var idRefs = ParserRuleContextHelper.GetDescendents(parentCtxt)
                .Where(desc => desc is LExprContext)
                    .Select(lexpr => GetTheSelectCaseReference(parentCtxt.Parent, lexpr.GetText()))
                    .Where(idr => idr != null);

            if (idRefs.Any())
            {
                //All typeNames of all ExpressionContexts are the same
                if(idRefs.All(idr => idRefs.First().Declaration.AsTypeName.Equals( idr.Declaration.AsTypeName)))
                {
                    selectStmtDO.IdReferenceName = idRefs.First().IdentifierName;
                    selectStmtDO =  SetAsAndBaseType(selectStmtDO, idRefs.First().Declaration.AsTypeName);
                    if (!idRefs.First().Declaration.AsTypeIsBaseType && idRefs.First().Declaration.AsTypeDeclaration.AsTypeIsBaseType)
                    {
                        var selectStmtTypeDecs = State.DeclarationFinder.MatchName(selectStmtDO.AsTypeName);
                        if(selectStmtTypeDecs.First().Context is EnumerationStmtContext)
                        {
                            var members = ParserRuleContextHelper.GetChildren<EnumerationStmt_ConstantContext>(selectStmtTypeDecs.First().Context);
                            foreach( var member in members)
                            {
                                var idCtxt = ParserRuleContextHelper.GetChild<IdentifierContext>(member);
                                var declarations = State.DeclarationFinder.MatchName(idCtxt.GetText());

                                var parentEval = new ExpressionEval
                                {
                                    IsUnaryOperation = true,
                                    Operator = "=",
                                    CanBeInspected = true,
                                    TypeNameTarget = Tokens.Long,
                                    SelectCaseRefName = idCtxt.GetText()
                                };

                                var contextEvals = AddEvaluationData(new Dictionary<ParserRuleContext, ExpressionEval>(), declarations.First().Context, parentEval);

                                var vbaValue = ResolveContextValue(contextEvals, declarations.First().Context).First().Value.Result;
                                selectStmtDO.SummaryCaseClauses.EnumDiscretes.Add(vbaValue);
                            }
                        }
                        selectStmtDO.BaseTypeName = idRefs.First().Declaration.AsTypeDeclaration.AsTypeName;
                    }

                    return selectStmtDO;
                }
                else
                {
                    //A mix of types in the Select Case statement
                    var variousTypes = idRefs.Select(idr => idr.Declaration.AsTypeName);
                    if (variousTypes.Contains(Tokens.String))
                    {
                        selectStmtDO.CanBeInspected = false;
                        return selectStmtDO;
                    }

                    if (variousTypes.Contains(Tokens.Currency))
                    {
                        return SetAsAndBaseType(selectStmtDO, Tokens.Currency);
                    }

                    if (variousTypes.Contains(Tokens.Double) || variousTypes.Contains(Tokens.Single))
                    {
                        return SetAsAndBaseType(selectStmtDO, Tokens.Double);
                    }

                    if (variousTypes.Contains(Tokens.Long)
                        || variousTypes.Contains(Tokens.LongLong)
                        || variousTypes.Contains(Tokens.Integer)
                        || variousTypes.Contains(Tokens.Byte))
                    {
                        return SetAsAndBaseType(selectStmtDO, Tokens.Long);
                    }
                }
            }
            return selectStmtDO;
        }

        private SelectStmtDataObject SetAsAndBaseType(SelectStmtDataObject selectStmtDO, string typeName)
        {
            selectStmtDO.AsTypeName = typeName;
            selectStmtDO.BaseTypeName = typeName;
            return selectStmtDO;
        }

        private string ResolveContextType(SelectCaseStmtContext selectExpr, LExprContext ctxt)
        {
            var idRef = GetTheSelectCaseReference(selectExpr, ctxt.GetText());
            return idRef != null ? idRef.Declaration.AsTypeName : string.Empty;
        }

        private IdentifierReference GetTheSelectCaseReference(RuleContext selectCaseStmtCtxt, string theName)
        {
            var identifierReferences = (State.DeclarationFinder.MatchName(theName).Select(dec => dec.References)).SelectMany(rf => rf);

            //TODO: Is there a scenario that results in two or more different references (same name within SelectStmtContext)?
            return identifierReferences.Any() ? identifierReferences.Where(idr => ParserRuleContextHelper.HasParent<SelectCaseStmtContext>(selectCaseStmtCtxt)).First() : null;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private SelectStmtDataObject InspectSelectStmtCaseClauses(SelectStmtDataObject selectStmtDO)
        {
            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count; idx++)
            {
                var caseClauseDO = selectStmtDO.CaseClauseDOs[idx];
                if (caseClauseDO.ResultType != ClauseEvaluationResult.Unreachable)
                {
                    for (var rgIdx = 0; rgIdx < caseClauseDO.RangeClauseDOs.Count; rgIdx++)
                    {
                        var rgClause = caseClauseDO.RangeClauseDOs[rgIdx];
                        rgClause.CanBeInspected = RangeClauseCanBeInspected(rgClause.Context, selectStmtDO.IdReferenceName);
                        if (rgClause.CanBeInspected)
                        {
                            rgClause = InitializeRangeClauseDataObject(rgClause, selectStmtDO.BaseTypeName, selectStmtDO.IdReferenceName);
                        }
                        else
                        {
                            rgClause.ResultType = ClauseEvaluationResult.NoResult;
                        }
                        caseClauseDO.RangeClauseDOs[rgIdx] = rgClause;
                    }
                }
                caseClauseDO.ResultType = caseClauseDO.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.MismatchType)
                    ? ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;
                selectStmtDO.CaseClauseDOs[idx] = caseClauseDO;
            }

            for (var idx = 0; idx < selectStmtDO.CaseClauseDOs.Count(); idx++)
            {
                var caseClause = selectStmtDO.CaseClauseDOs[idx];
                if (selectStmtDO.SummaryCaseClauses.UnreachableCaseElse)
                {
                    caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                }
                else
                {
                    selectStmtDO = InspectCaseClause(ref caseClause, selectStmtDO);

                    //Final inspection to look for copy/paste duplicate Cases that resist other analyses
                    if(caseClause.ResultType == ClauseEvaluationResult.NoResult)
                    {
                        if (caseClause.RangeClauseDOs.All(rg => selectStmtDO.SummaryCaseClauses.RangeClausesAsText.Contains(rg.AsText)))
                        {
                            caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                        }
                    }
                    selectStmtDO.SummaryCaseClauses.RangeClausesAsText.AddRange(caseClause.RangeClauseDOs.Select(rg => rg.AsText));
                }
                selectStmtDO.CaseClauseDOs[idx] = caseClause;
            }
            return selectStmtDO;
        }

        private SelectStmtDataObject InspectCaseClause(ref CaseClauseDataObject caseClause, SelectStmtDataObject selectStmtDO)
        {
            if (caseClause.ResultType != ClauseEvaluationResult.NoResult)
            {
                return selectStmtDO;
            }

            for (var idx = 0; idx < caseClause.RangeClauseDOs.Count(); idx++)
            {
                var rangeClauseDO = caseClause.RangeClauseDOs[idx];
                if (rangeClauseDO.ResultType != ClauseEvaluationResult.NoResult && rangeClauseDO.CanBeInspected)
                {
                    continue;
                }

                if (rangeClauseDO.IsValueRange)
                {
                    rangeClauseDO = InspectValueRangeClause(selectStmtDO, rangeClauseDO);
                    selectStmtDO.SummaryCaseClauses = UpdateSummaryDataRanges(selectStmtDO.SummaryCaseClauses, rangeClauseDO);
                }
                else
                {
                    rangeClauseDO = InspectSingleValueRangeClause(selectStmtDO, rangeClauseDO);
                    selectStmtDO.SummaryCaseClauses = UpdateSummaryDataSingleValues(selectStmtDO.SummaryCaseClauses, rangeClauseDO);
                }
                caseClause.RangeClauseDOs[idx] = rangeClauseDO;
            }

            caseClause.ResultType = caseClause.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.Unreachable)
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            return selectStmtDO;
        }

        private RangeClauseDataObject InspectValueRangeClause(SelectStmtDataObject selectStmtDO, RangeClauseDataObject rangeClauseDO)
        {
            if (rangeClauseDO.MinValue == null || rangeClauseDO.MaxValue == null)
            {
                return rangeClauseDO;
            }

            if (rangeClauseDO.MinValue.ExceedsMaxMin() && rangeClauseDO.MaxValue.ExceedsMaxMin())
            {
                rangeClauseDO.ResultType = ClauseEvaluationResult.Unreachable;
                return rangeClauseDO;
            }

            var minValue = rangeClauseDO.MinValue;
            var maxValue = rangeClauseDO.MaxValue;
            if (selectStmtDO.SummaryCaseClauses.EnumDiscretes.Any())
            {
                var rangeContainsAnEnumValue = selectStmtDO.SummaryCaseClauses.EnumDiscretes.Any(env => env.IsWithin(minValue, maxValue));

                rangeClauseDO.ResultType = !rangeContainsAnEnumValue
                        ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
                if (rangeClauseDO.ResultType == ClauseEvaluationResult.NoResult)
                {
                    rangeClauseDO.ResultType = selectStmtDO.SummaryCaseClauses.Ranges.Where(rg => minValue.IsWithin(rg.Item1, rg.Item2)
                           && maxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                           || selectStmtDO.SummaryCaseClauses.IsLTMax != null && selectStmtDO.SummaryCaseClauses.IsLTMax > rangeClauseDO.MaxValue
                           || selectStmtDO.SummaryCaseClauses.IsGTMin != null && selectStmtDO.SummaryCaseClauses.IsGTMin < rangeClauseDO.MinValue
                           ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
                }
            }
            else
            {
                rangeClauseDO.ResultType = selectStmtDO.SummaryCaseClauses.Ranges.Where(rg => minValue.IsWithin(rg.Item1, rg.Item2)
                        && maxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                        || selectStmtDO.SummaryCaseClauses.IsLTMax != null && selectStmtDO.SummaryCaseClauses.IsLTMax > rangeClauseDO.MaxValue
                        || selectStmtDO.SummaryCaseClauses.IsGTMin != null && selectStmtDO.SummaryCaseClauses.IsGTMin < rangeClauseDO.MinValue
                        ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
            }

            return rangeClauseDO;
        }

        private RangeClauseDataObject InspectSingleValueRangeClause(SelectStmtDataObject selectStmtDO, RangeClauseDataObject rangeClauseDO)
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

            if (rangeClauseDO.UsesIsClause)
            {
                if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.LT)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.LTE))
                {
                    rangeClauseDO.ResultType = selectStmtDO.SummaryCaseClauses.IsLTMax != null && selectStmtDO.SummaryCaseClauses.IsLTMax >= rangeClauseDO.SingleValue
                            ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
                }
                else if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.GT)
                        || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GTE))
                {
                    rangeClauseDO.ResultType = selectStmtDO.SummaryCaseClauses.IsGTMin != null && selectStmtDO.SummaryCaseClauses.IsGTMin <= rangeClauseDO.SingleValue
                            ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
                }
                else if (CompareTokens.EQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    rangeClauseDO.ResultType = HandleSimpleSingleValueCompare(rangeClauseDO.SingleValue, selectStmtDO);
                }
                else if (CompareTokens.NEQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    if (selectStmtDO.SummaryCaseClauses.SingleValues.Contains(rangeClauseDO.SingleValue))
                    {
                        rangeClauseDO.ResultType = ClauseEvaluationResult.CaseElse;
                    }
                }
            }
            else
            {
                rangeClauseDO.ResultType = HandleSimpleSingleValueCompare(rangeClauseDO.SingleValue, selectStmtDO);
            }
            return rangeClauseDO;
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

        private bool SingleValueIsHandledPreviously(VBAValue theValue, SummaryCaseCoverage priorHandlers)
        {
            if (theValue.OriginTypeName.Equals(Tokens.Boolean))
            {
                //TODO: do we need to do more?
                return priorHandlers.SingleValues.Any(val => val.AsBoolean() == theValue.AsBoolean());
            }
            else
            {
                return priorHandlers.IsLTMax != null && theValue < priorHandlers.IsLTMax
                    || priorHandlers.IsGTMin != null && theValue > priorHandlers.IsGTMin
                    || priorHandlers.SingleValues.Contains(theValue)
                    || priorHandlers.Ranges.Where(rg => theValue.IsWithin(rg.Item1, rg.Item2)).Any();
            }
        }

        private SummaryCaseCoverage UpdateSummaryIsClauseLimits(VBAValue theValue, string compareSymbol, SummaryCaseCoverage priorHandlers)
        {
            if (compareSymbol.Equals(CompareTokens.LT) || compareSymbol.Equals(CompareTokens.LTE) ) // new string[] { CompareSymbols.LT, CompareSymbols.LTE }.Contains(compareSymbol))
            {
                priorHandlers.IsLTMax = priorHandlers.IsLTMax == null ? theValue
                    : priorHandlers.IsLTMax < theValue ? theValue : priorHandlers.IsLTMax;
            }
            else if (compareSymbol.Equals(CompareTokens.GT) || compareSymbol.Equals(CompareTokens.GTE))
            {
                priorHandlers.IsGTMin = priorHandlers.IsGTMin == null ? theValue
                    : priorHandlers.IsGTMin > theValue ? theValue : priorHandlers.IsGTMin;
            }
            else
            {
                return priorHandlers;
            }

            if (CompareTokens.LTE == compareSymbol || CompareTokens.GTE == compareSymbol)
            {
                if (!priorHandlers.SingleValues.Contains(theValue))
                {
                    priorHandlers.SingleValues.Add(theValue);
                }
            }
            return priorHandlers;
        }

        private SelectStmtDataObject HandleSimpleSingleValueCompare(ref RangeClauseDataObject range, VBAValue theValue, SelectStmtDataObject selectStmtDO)
        {
            range.ResultType = SingleValueIsHandledPreviously(theValue, selectStmtDO.SummaryCaseClauses)
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            if (theValue.OriginTypeName.Equals(Tokens.Boolean))
            {
                range.ResultType = selectStmtDO.SummaryCaseClauses.SingleValues.Any() && selectStmtDO.SummaryCaseClauses.SingleValues.Any(val => val.AsBoolean() != theValue.AsBoolean())
                    ? ClauseEvaluationResult.CaseElse : range.ResultType;
            }
            return selectStmtDO;
        }

        private ClauseEvaluationResult HandleSimpleSingleValueCompare(VBAValue theValue, SelectStmtDataObject selectStmtDO)
        {
            var result = SingleValueIsHandledPreviously(theValue, selectStmtDO.SummaryCaseClauses)
                ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            if(result == ClauseEvaluationResult.NoResult)
            {
                if (selectStmtDO.SummaryCaseClauses.EnumDiscretes.Any())
                {
                    result = !selectStmtDO.SummaryCaseClauses.EnumDiscretes.Contains(theValue) ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
                }

                if (theValue.OriginTypeName.Equals(Tokens.Boolean))
                {
                    result = selectStmtDO.SummaryCaseClauses.SingleValues.Any() && selectStmtDO.SummaryCaseClauses.SingleValues.Any(val => val.AsBoolean() != theValue.AsBoolean())
                        ? ClauseEvaluationResult.CaseElse : result;
                }
            }
            return result;
        }

        private SummaryCaseCoverage UpdateSummaryData(SummaryCaseCoverage summaryClauses, RangeClauseDataObject rangeClauseDO)
        {
            if(!rangeClauseDO.CanBeInspected) { return summaryClauses; }

            if (rangeClauseDO.IsValueRange)
            {
                return summaryClauses = UpdateSummaryDataRanges(summaryClauses, rangeClauseDO);
            }
            else
            {
                return summaryClauses = UpdateSummaryDataSingleValues(summaryClauses, rangeClauseDO);
            }
        }

        private SummaryCaseCoverage UpdateSummaryDataSingleValues(SummaryCaseCoverage summaryClauses, RangeClauseDataObject rangeClauseDO)
        {
            if (!rangeClauseDO.CanBeInspected || rangeClauseDO.SingleValue == null) { return summaryClauses; }

            //if (rangeClauseDO.ResultType != ClauseEvaluationResult.NoResult)
            //{
            //    return summaryClauses;
            //}

            if (rangeClauseDO.UsesIsClause)
            {
                if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.LT)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.LTE)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GT)
                    || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GTE)
                    )
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

            if (!summaryClauses.UnreachableCaseElse)
            {
                summaryClauses.UnreachableCaseElse = IsClausesCoverAllValues(summaryClauses)
                    || (rangeClauseDO.TypeNameTarget.Equals(Tokens.Boolean)
                    && summaryClauses.SingleValues.Any(sv => sv.AsLong().Value != 0) && summaryClauses.SingleValues.Any(sv => sv.AsLong().Value == 0));
            }
            return summaryClauses;
        }

        private SummaryCaseCoverage UpdateSummaryDataRanges(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO)
        {
            if (!rangeClauseDO.CanBeInspected || !rangeClauseDO.IsValueRange) { return summaryCoverage; }

            if (rangeClauseDO.TypeNameTarget.Equals(Tokens.Boolean))
            {
                for( var theVal = rangeClauseDO.MinValue.AsLong().Value; theVal <= rangeClauseDO.MaxValue.AsLong().Value; theVal++)
                {
                    summaryCoverage.SingleValues.Add(new VBAValue(theVal, Tokens.Long));
                }
            }
            else if (summaryCoverage.EnumDiscretes.Any())
            {
                var used = summaryCoverage.EnumDiscretes.Where(ed => ed.IsWithin(rangeClauseDO.MinValue, rangeClauseDO.MaxValue));
                summaryCoverage.SingleValues.AddRange(used);
            }


            var updatedRanges = new List<Tuple<VBAValue, VBAValue>>();
            var overlapsMin = summaryCoverage.Ranges.Where(rg => rangeClauseDO.MinValue.IsWithin(rg.Item1, rg.Item2));
            var overlapsMax = summaryCoverage.Ranges.Where(rg => rangeClauseDO.MaxValue.IsWithin(rg.Item1, rg.Item2));
            foreach (var rg in summaryCoverage.Ranges)
            {
                if (overlapsMin.Contains(rg))
                {
                    updatedRanges.Add(new Tuple<VBAValue, VBAValue>(rg.Item1, rangeClauseDO.MaxValue));
                }
                else if (overlapsMax.Contains(rg))
                {
                    updatedRanges.Add(new Tuple<VBAValue, VBAValue>(rangeClauseDO.MinValue, rg.Item2));
                }
                else
                {
                    updatedRanges.Add(rg);
                }
            }

            if (!overlapsMin.Any() && !overlapsMax.Any())
            {
                updatedRanges.Add(new Tuple<VBAValue, VBAValue>(rangeClauseDO.MinValue, rangeClauseDO.MaxValue));
            }

            summaryCoverage.Ranges = updatedRanges;

            summaryCoverage = AggregateSummaryRanges(summaryCoverage);

            return summaryCoverage;
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

        private List<Tuple<VBAValue, VBAValue>> AppendRanges(List<Tuple<VBAValue, VBAValue>> ranges)
        {
            if (ranges.Count() <= 1 || !IsIntegerNumberType(ranges.First().Item1.UseageTypeName))
            {
                return ranges;
            }

            var updatedRanges = new List<Tuple<VBAValue, VBAValue>>();
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
                    updatedRanges.Add(new Tuple<VBAValue, VBAValue>(theMin, theNextMax));
                    combinedLastRange = true;
                }
                else if (theMin.AsLong() == theNextMax.AsLong() + 1)
                {
                    updatedRanges.Add(new Tuple<VBAValue, VBAValue>(theNextMin, theMax));
                    combinedLastRange = true;
                }
                else
                {
                    updatedRanges.Add(ranges[idx]);
                }
            }
            return updatedRanges;
        }

        private static bool IsIntegerNumberType(string typeName) => new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte }.Contains(typeName);

        private static bool HasChildToken<T>(T ctxt, string token) where T : ParserRuleContext
        {
            return ctxt.children.Any(ch => ch.GetText().Equals(token));
        }

        private RangeClauseDataObject SetTheCompareOperator(RangeClauseDataObject rangeClauseDO)
        {
            rangeClauseDO.UsesIsClause = TryGetChildContext(rangeClauseDO.Context, out ComparisonOperatorContext opCtxt);
            rangeClauseDO.CompareSymbol = rangeClauseDO.UsesIsClause ? opCtxt.GetText() : CompareTokens.EQ;
            return rangeClauseDO;
        }

        private static bool TryGetChildContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : ParserRuleContext
        {
            opCtxt = ParserRuleContextHelper.GetChild<U>(ctxt);
            return opCtxt != null;
        }

        private static string GetText(ParserRuleContext ctxt) => ctxt.GetText().Replace("\"", "");

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

        private bool RangeClauseCanBeInspected(RangeClauseContext context, string refName)
        {
            var canBeInspected = true;
            var ipts = ParserRuleContextHelper.GetDescendents(context);
            var ops = ipts.Where(ipt => (ipt is ParserRuleContext) && (IsBinaryMathOperation(ipt) || IsUnaryMathOperation(ipt)));
            foreach (var op in ops)
            {
                var lExprCtxts = ParserRuleContextHelper.GetChildren<LExprContext>((RuleContext)op);
                if (lExprCtxts.Any(lex => lex.GetText().Equals(refName)))
                {
                    //TODO: (Future)implement the necessary algebra to support these use cases
                    canBeInspected = false;
                }
            }
            return canBeInspected;
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