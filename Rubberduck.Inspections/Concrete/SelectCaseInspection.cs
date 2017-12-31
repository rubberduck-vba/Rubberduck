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
        public static string MismatchType => "Type cannot be converted to the Select Statement Type";
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
            public bool CaseElseIsUnreachable;
            public List<string> RangeClausesAsText;
            public List<VBAValue> EnumDiscretes;
        }

        internal struct ExpressionEvaluationDataObject
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

            public SelectStmtDataObject(QualifiedContext<ParserRuleContext> selectStmtCtxt)
            {
                SelectStmtContext = (SelectCaseStmtContext)selectStmtCtxt.Context;
                IdReferenceName = string.Empty;
                BaseTypeName = Tokens.Variant;
                AsTypeName = Tokens.Variant;
                CaseClauseDOs = new List<CaseClauseDataObject>();
                CaseElseContext = ParserRuleContextHelper.GetChild<VBAParser.CaseElseClauseContext>(SelectStmtContext);
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

                selectStmtDO = InitializeCaseClauses(selectStmtDO);

                if (!selectStmtDO.CanBeInspected) { continue; }

                selectStmtDO = InspectSelectStmtCaseClauses(selectStmtDO);

                if (!selectStmtDO.CanBeInspected) { continue; }

                inspResults.AddRange(selectStmtDO.CaseClauseDOs.Where(cc => cc.ResultType != ClauseEvaluationResult.NoResult)
                    .Select(cc => CreateInspectionResult(selectStmt, cc.CaseContext, _resultMessages[cc.ResultType])));

                if (selectStmtDO.SummaryCaseClauses.CaseElseIsUnreachable && selectStmtDO.CaseElseContext != null)
                {
                    inspResults.Add(CreateInspectionResult(selectStmt, selectStmtDO.CaseElseContext, _resultMessages[ClauseEvaluationResult.CaseElse]));
                }
            }
            return inspResults;
        }

        private SelectStmtDataObject InitializeSelectStatementDataObject(SelectStmtDataObject selectStmtDO)
        {
            selectStmtDO = ResolveSelectStmtControlVariableNameAndType(selectStmtDO);

            if (selectStmtDO.CanBeInspected && selectStmtDO.BaseTypeName.Equals(Tokens.Variant))
            {
                if (TryInferUseageTypeForUnResolvedType( selectStmtDO, out string typeName))
                {
                    selectStmtDO = UpdateTypeNames(selectStmtDO, typeName);
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

        private bool TryInferUseageTypeForUnResolvedType(SelectStmtDataObject selectStmtDO, out string typeName)
        {
            typeName = selectStmtDO.BaseTypeName;

            if (!selectStmtDO.BaseTypeName.Equals(Tokens.Variant)) { return false; }

            var rangeCtxts = ParserRuleContextHelper.GetChildren<CaseClauseContext>(selectStmtDO.SelectStmtContext)
                .Select(cc => ParserRuleContextHelper.GetChildren<RangeClauseContext>(cc))
                .SelectMany(rgCtxt => rgCtxt);

            var typeNames = rangeCtxts.SelectMany(rgCtxt => ParserRuleContextHelper.GetDescendents(rgCtxt))
                .Where(desc => desc is LiteralExprContext || desc is LExprContext)
                    .Select(exprCtxt => EvaluateContextTypeName((ExpressionContext)exprCtxt, selectStmtDO))
                    .Where(tn => tn != string.Empty);

            if (typeNames.All(tn => typeNames.First().Equals(tn))) //they all match
            {
                typeName = typeNames.First();
                return true;
            }

            //All cases can be evaluated using Long
            if (typeNames.All(tn => tn.Equals(Tokens.Long)
                    || tn.Equals(Tokens.LongLong)
                    || tn.Equals(Tokens.Integer)
                    || tn.Equals(Tokens.Byte)))
            {
                typeName = Tokens.Long;
                return true;
            }
            //All cases can be evaluated using Double
            if (typeNames.All(tn => !(tn.Equals(Tokens.Currency) || tn.Equals(Tokens.String))))
            {
                typeName = Tokens.Double;
                return true;
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
                var startContext = ParserRuleContextHelper.GetChild<SelectStartValueContext>(rangeClauseDO.Context);
                var endContext = ParserRuleContextHelper.GetChild<SelectEndValueContext>(rangeClauseDO.Context);
                rangeClauseDO = ResolveRangeClauseValue(rangeClauseDO, startContext, out VBAValue startValue);
                rangeClauseDO = ResolveRangeClauseValue(rangeClauseDO, endContext, out VBAValue endValue);

                var startTypeName = startValue.HasValue ? startValue.UseageTypeName : startValue.DerivedTypeName;
                var endTypeName = endValue.HasValue ? endValue.UseageTypeName : endValue.DerivedTypeName;

                if (!startTypeName.Equals(endTypeName))
                {
                    var typePrecedence = new string[] { Tokens.Double, Tokens.Long, Tokens.Integer, Tokens.Byte };
                    //Find acceptable type for comparisons if possible
                    if (typePrecedence.Contains(startTypeName) && typePrecedence.Contains(endTypeName))
                    {
                        for (var idx = 0; idx < typePrecedence.Count(); idx++)
                        {
                            if (typePrecedence[idx].Equals(startTypeName) || typePrecedence[idx].Equals(endTypeName))
                            {
                                startValue = new VBAValue(startValue.AsString(), typePrecedence[idx]);
                                endValue = new VBAValue(endValue.AsString(), typePrecedence[idx]);
                                rangeClauseDO.TypeNameDerived = typePrecedence[idx];
                                idx = typePrecedence.Count();
                            }
                        }
                    }
                    else
                    {
                        rangeClauseDO.TypeNameDerived = string.Empty;
                        rangeClauseDO.ResultType = ClauseEvaluationResult.MismatchType;
                        return rangeClauseDO;
                    }
                }
                if(startValue != null && endValue != null)
                {
                    rangeClauseDO.MinValue = startValue <= endValue ? startValue : endValue;
                    rangeClauseDO.MaxValue = startValue <= endValue ? endValue : startValue;
                    rangeClauseDO.SingleValue = rangeClauseDO.MinValue;
                    rangeClauseDO.IsParseable = rangeClauseDO.MinValue.HasValue && rangeClauseDO.MaxValue.HasValue;
                    rangeClauseDO.CompareByTextOnly = !rangeClauseDO.IsParseable && rangeClauseDO.TypeNameDerived.Equals(rangeClauseDO.TypeNameTarget);
                    rangeClauseDO.ResultType = !rangeClauseDO.TypeNameTarget.Equals(rangeClauseDO.TypeNameDerived) && !rangeClauseDO.IsParseable ?
                            ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;
                }
                else
                {
                    rangeClauseDO.IsParseable = false;
                    rangeClauseDO.CompareByTextOnly = true;
                    rangeClauseDO.ResultType = ClauseEvaluationResult.NoResult;
                }
            }
            else
            {
                rangeClauseDO = ResolveRangeClauseValue(rangeClauseDO, rangeClauseDO.Context, out VBAValue value);
                if (value != null)
                {
                    rangeClauseDO.MinValue = value;
                    rangeClauseDO.MaxValue = value;
                    rangeClauseDO.SingleValue = value;
                    rangeClauseDO.TypeNameDerived = rangeClauseDO.SingleValue.DerivedTypeName;
                    rangeClauseDO.IsParseable = rangeClauseDO.SingleValue == null ? false : rangeClauseDO.SingleValue.HasValue;
                    rangeClauseDO.CompareByTextOnly = !rangeClauseDO.IsParseable && rangeClauseDO.TypeNameDerived.Equals(rangeClauseDO.TypeNameTarget);
                    rangeClauseDO.ResultType = !rangeClauseDO.TypeNameTarget.Equals(rangeClauseDO.TypeNameDerived) && !rangeClauseDO.IsParseable ?
                            ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;
                    rangeClauseDO.CanBeInspected = rangeClauseDO.IsParseable;
                }
                else
                {
                    rangeClauseDO.IsParseable = false;
                    rangeClauseDO.CompareByTextOnly = true;
                    rangeClauseDO.ResultType = ClauseEvaluationResult.NoResult;
                    rangeClauseDO.CanBeInspected = rangeClauseDO.IsParseable;
                }
            }
            return rangeClauseDO;
        }

        private RangeClauseDataObject ResolveRangeClauseValue(RangeClauseDataObject rangeClauseDO, ParserRuleContext context, out VBAValue vbaValue)
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

            var contextEvals = AddEvaluationData(new Dictionary<ParserRuleContext, ExpressionEvaluationDataObject>(), context, parentEval);

            contextEvals = ResolveContextValue(contextEvals, context);
            rangeClauseDO.CompareSymbol = contextEvals[context].Operator;
            rangeClauseDO.UsesIsClause = rangeClauseDO.UsesIsClause ? rangeClauseDO.UsesIsClause : contextEvals[context].EvaluateAsIsClause;
            vbaValue =  contextEvals[context].Result;
            return rangeClauseDO;
        }

        private Dictionary<ParserRuleContext,ExpressionEvaluationDataObject> ResolveContextValue( Dictionary<ParserRuleContext, ExpressionEvaluationDataObject> contextEvals, ParserRuleContext parentContext)
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
                    childData.LHSValue = EvaluateContextValue((ExpressionContext)child, childData.TypeNameTarget);
                    childData.Result = childData.LHSValue;

                    contextEvals = AddEvaluationData(contextEvals, (ParserRuleContext)child, childData);
                    contextEvals = UpdateParentEvaluation((ParserRuleContext)child, contextEvals);
                }
            }
            return contextEvals;
        }

        private Dictionary<ParserRuleContext, ExpressionEvaluationDataObject> UpdateParentEvaluation(ParserRuleContext child, Dictionary<ParserRuleContext, ExpressionEvaluationDataObject> ctxtEvalResults)
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
                parentData.Result = childData.ParentCtxt is UnaryMinusOpContext ? 
                    parentData.LHSValue.AdditiveInverse : parentData.LHSValue;
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
                    if (!parentData.RHSValue.HasValue && !parentData.RHSValue.AsString().Equals(childData.SelectCaseRefName))
                    {
                        childData.CanBeInspected = false;
                        parentData.CanBeInspected = false;
                    }
                    else
                    {
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
            }
            return AddEvaluationData(ctxtEvalResults, childData.ParentCtxt, parentData);
        }

        private ExpressionEvaluationDataObject GetEvaluationData(ParserRuleContext ctxt, Dictionary<ParserRuleContext, ExpressionEvaluationDataObject> ctxtEvalResults)
        {
            return ctxtEvalResults.ContainsKey(ctxt) ? ctxtEvalResults[ctxt] : new ExpressionEvaluationDataObject { Operator = string.Empty, CanBeInspected = true };
        }

        private Dictionary<ParserRuleContext, ExpressionEvaluationDataObject> AddEvaluationData(Dictionary<ParserRuleContext, ExpressionEvaluationDataObject> contextIndices, ParserRuleContext ctxt, ExpressionEvaluationDataObject exprEvaluation)
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

        private string EvaluateContextTypeName(ExpressionContext ctxt, SelectStmtDataObject selectStmtDO)
        {
            var val = EvaluateContextValue(ctxt, selectStmtDO.BaseTypeName);
            return val != null ? val.DerivedTypeName : string.Empty;
        }

        private VBAValue EvaluateContextValue(ExpressionContext ctxt, string typeName)
        {
            if (ctxt is LExprContext)
            {
                if(TryGetTheExpressionValue((LExprContext)ctxt, out string lexpr))
                {
                    return new VBAValue(lexpr, typeName);
                }
                var identifierReferences = (State.DeclarationFinder.MatchName(ctxt.GetText()).Select(dec => dec.References)).SelectMany(rf => rf);
                var idRefs = identifierReferences.Where(idr => idr.Context.Parent == ctxt);
                if (idRefs.Any())
                {
                    var theTypeName = GetBaseTypeForDeclaration(idRefs.First().Declaration);
                    return new VBAValue(ctxt.GetText(), theTypeName);
                }
                else
                {
                    return new VBAValue(ctxt.GetText(), typeName);
                }
            }
            else if (ctxt is LiteralExprContext)
            {
                return new VBAValue(ctxt.GetText(), typeName);
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
                //is the SelectCase variable.  The result is '4' and
                //is handled as if it was an 'Is' statement e.g., 'Is < 4'
                return LHS.HasValue ? CompareOperations[operation](LHS, RHS).AsString() : RHS.AsString();
            }
            return string.Empty;
        }

        private bool TryGetTheExpressionValue(LExprContext ctxt, out string expressionValue)
        {
            expressionValue = string.Empty;
            var expressionType = string.Empty;
            if(TryGetChildContext(ctxt, out MemberAccessExprContext member))
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
                            expressionType = dec.AsTypeIsBaseType ? dec.AsTypeName : dec.AsTypeDeclaration.AsTypeName;
                            return true;
                        }
                    }
                }
                return false;
            }
            else if (TryGetChildContext(ctxt, out SimpleNameExprContext smplName))
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
                    var parentData = new Dictionary<ParserRuleContext, ExpressionEvaluationDataObject>();
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
            var eqIndex = contexts.FindIndex(ch => ch.GetText().Equals(CompareTokens.EQ));
            if(eqIndex == contexts.Count)
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

        private SelectStmtDataObject ResolveSelectStmtControlVariableNameAndType(SelectStmtDataObject selectStmtDO )
        {
            var selectExprCtxt = selectStmtDO.SelectExpressionContext;
            if (selectExprCtxt.children.Count != 1)
            {
                selectStmtDO.CanBeInspected = false;
                return selectStmtDO;
            }

            var child = selectExprCtxt.children.First();
            if (IsBinaryLogicalOperation(child) || IsUnaryLogicalOperator(child))
            {
                return UpdateTypeNames(selectStmtDO, Tokens.Boolean);
            }

            var selectLExprs = ParserRuleContextHelper.GetDescendents(selectExprCtxt)
                .Where(desc => desc is LExprContext)
                .Select(desc => desc as LExprContext);

            var selectLExprCandidates = new Dictionary<string, Tuple<string, string>>();
            foreach (var selectLExpr in selectLExprs)
            {
                var identifierReferences = (State.DeclarationFinder.MatchName(selectLExpr.GetText()).Select(dec => dec.References)).SelectMany(rf => rf);
                var idRefs = identifierReferences.Where(idr => idr.Context.Parent == selectLExpr);
                {
                    foreach (var idRef in idRefs)
                    {
                        var asType = idRef.Declaration.AsTypeName;
                        var ctxtVal = EvaluateContextValue(selectLExpr, string.Empty);
                        if (!ctxtVal.HasValue && ctxtVal.OriginTypeName.Length > 0)
                        {
                            selectLExprCandidates.Add(idRef.IdentifierName, new Tuple<string, string>(idRef.Declaration.AsTypeName, GetBaseTypeForDeclaration(idRef.Declaration)));
                        }
                    }
                }
            }

            if (selectLExprCandidates.Keys.Count == 0)
            {
                string typeName = string.Empty;
                if (TryInferUseageTypeForUnResolvedType(selectStmtDO, out typeName))
                {
                    return UpdateTypeNames(selectStmtDO, typeName);
                }
            }
            else if (selectLExprCandidates.Keys.Count == 1)
            {
                return UpdateRefNameAndTypes(selectStmtDO, selectLExprCandidates.First().Key, selectLExprCandidates.First().Value.Item1, selectLExprCandidates.First().Value.Item2);
            }

            else if (selectLExprCandidates.Keys.Count > 1)  // e.g. Select Case x * y (where neither x nor y are constants
            {
                var typePrecedence = new string[] { Tokens.Double, Tokens.Long, Tokens.Integer, Tokens.Byte };
                if(selectLExprCandidates.Values.All(vs => selectLExprCandidates.Values.First().Item2 == vs.Item2))
                {
                    return UpdateRefNameAndTypes(selectStmtDO, selectLExprCandidates.First().Key, selectLExprCandidates.First().Value.Item1, selectLExprCandidates.First().Value.Item2);
                }
                else if(selectLExprCandidates.Values.All(vs => typePrecedence.Contains(vs.Item2)))
                {
                    foreach(var type in typePrecedence)
                    {
                        foreach (var candidate in selectLExprCandidates)
                        {
                            if (candidate.Value.Item2 == type)
                            {
                                return UpdateRefNameAndTypes(selectStmtDO, candidate.Key, candidate.Value.Item1, candidate.Value.Item2);
                            }
                        }
                    }
                    return selectStmtDO;
                }
            }

            selectStmtDO.CanBeInspected = false;
            return selectStmtDO;
        }

        private SelectStmtDataObject UpdateRefNameAndTypes(SelectStmtDataObject selectStmtDO, string idRefName, string asType, string baseType)
        {
            selectStmtDO.IdReferenceName = idRefName;
            selectStmtDO = UpdateTypeNames(selectStmtDO, asType, baseType);
            if (selectStmtDO.BaseTypeName != selectStmtDO.AsTypeName)
            {
                return InitializeEnumSummaryData(selectStmtDO);
            }
            return selectStmtDO;
        }

        private SelectStmtDataObject InitializeEnumSummaryData(SelectStmtDataObject selectStmtDO)
        {
            var selectStmtTypeDecs = State.DeclarationFinder.MatchName(selectStmtDO.AsTypeName).Where(dec => dec.Context is EnumerationStmtContext);
            if (selectStmtTypeDecs.Any() && selectStmtTypeDecs.Count() == 1)
            {
                var members = ParserRuleContextHelper.GetChildren<EnumerationStmt_ConstantContext>(selectStmtTypeDecs.First().Context);
                foreach (var member in members)
                {
                    var idCtxt = ParserRuleContextHelper.GetChild<IdentifierContext>(member);
                    var declarations = State.DeclarationFinder.MatchName(idCtxt.GetText());

                    var parentEval = new ExpressionEvaluationDataObject
                    {
                        IsUnaryOperation = true,
                        Operator = CompareTokens.EQ,
                        CanBeInspected = true,
                        TypeNameTarget = Tokens.Long,
                        SelectCaseRefName = idCtxt.GetText()
                    };

                    var contextEvals = AddEvaluationData(new Dictionary<ParserRuleContext, ExpressionEvaluationDataObject>(), declarations.First().Context, parentEval);

                    var vbaValue = ResolveContextValue(contextEvals, declarations.First().Context).First().Value.Result;
                    selectStmtDO.SummaryCaseClauses.EnumDiscretes.Add(vbaValue);
                }
                return selectStmtDO;
            }
            else
            {
                selectStmtDO.CanBeInspected = false;
                return selectStmtDO;
            }
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
            return identifierReferences.Any() ? identifierReferences.Where(idr => ParserRuleContextHelper.HasParent<SelectCaseStmtContext>(parentContext)).First() : null;
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
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
                        rgClause.CanBeInspected = RangeClauseCanBeInspected(rgClause.Context, selectStmtDO.IdReferenceName);
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

                //Once the CaseElse (whether CaseElse exists or not) is found unreachable, every remaining Case Clause is also unreachable
                if (selectStmtDO.SummaryCaseClauses.CaseElseIsUnreachable)
                {
                    caseClause.ResultType = ClauseEvaluationResult.Unreachable;
                    selectStmtDO.CaseClauseDOs[idx] = caseClause;
                    continue;
                }

                //Inspect for duplicate Case Clauses to short circuit the more costly analysis
                caseClause.ResultType = caseClause.RangeClauseDOs.All(rg => selectStmtDO.SummaryCaseClauses.RangeClausesAsText.Contains(rg.AsText)) ?
                    ClauseEvaluationResult.Unreachable : caseClause.ResultType;

                if(caseClause.ResultType == ClauseEvaluationResult.NoResult)
                {
                    caseClause = InspectCaseClause(caseClause, selectStmtDO.SummaryCaseClauses);
                    selectStmtDO = UpdateSummaryClauses(selectStmtDO, caseClause);

                    if (caseClause.RangeClauseDOs.All(rg => rg.ResultType != ClauseEvaluationResult.NoResult))
                    {
                        caseClause.ResultType = caseClause.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.Unreachable)
                            ? ClauseEvaluationResult.Unreachable : caseClause.RangeClauseDOs.All(rg => rg.ResultType == ClauseEvaluationResult.MismatchType)
                                ? ClauseEvaluationResult.MismatchType : ClauseEvaluationResult.NoResult;
                    }
                }
                selectStmtDO.CaseClauseDOs[idx] = caseClause;
            }
            return selectStmtDO;
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
            if (rangeClauseDO.MinValue == null || rangeClauseDO.MaxValue == null)
            {
                return rangeClauseDO;
            }

            if (rangeClauseDO.MinValue.ExceedsMaxMin() && rangeClauseDO.MaxValue.ExceedsMaxMin())
            {
                rangeClauseDO.ResultType = ClauseEvaluationResult.Unreachable;
                return rangeClauseDO;
            }

            if (summaryCoverage.EnumDiscretes.Any())
            {
                var capturedEnumValues = summaryCoverage.EnumDiscretes.Where(env => env.IsWithin(rangeClauseDO.MinValue, rangeClauseDO.MaxValue));

                rangeClauseDO.ResultType = !capturedEnumValues.Any() || capturedEnumValues.All(ev => summaryCoverage.SingleValues.Contains(ev))
                        ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;

            }

            if (rangeClauseDO.ResultType == ClauseEvaluationResult.NoResult)
            {
                rangeClauseDO.ResultType = summaryCoverage.Ranges.Where(rg => rangeClauseDO.MinValue.IsWithin(rg.Item1, rg.Item2)
                                   && rangeClauseDO.MaxValue.IsWithin(rg.Item1, rg.Item2)).Any()
                                    || summaryCoverage.IsLTMax != null && summaryCoverage.IsLTMax > rangeClauseDO.MaxValue
                                    || summaryCoverage.IsGTMin != null && summaryCoverage.IsGTMin < rangeClauseDO.MinValue
                                    ? ClauseEvaluationResult.Unreachable : ClauseEvaluationResult.NoResult;
            }

            return rangeClauseDO;
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
                    if (summaryCoverage.EnumDiscretes.Any())
                    {
                        var capturedEnums = rangeClauseDO.CompareSymbol.Equals(CompareTokens.LT) ?
                            summaryCoverage.EnumDiscretes.Where(en => en < rangeClauseDO.SingleValue)
                            : summaryCoverage.EnumDiscretes.Where(en => en <= rangeClauseDO.SingleValue);

                        isUnreachable = capturedEnums.All(en => summaryCoverage.SingleValues.Contains(en));
                    }
                    else
                    {
                        isUnreachable = summaryCoverage.IsLTMax != null && summaryCoverage.IsLTMax >= rangeClauseDO.SingleValue;
                    }
                }
                else if (rangeClauseDO.CompareSymbol.Equals(CompareTokens.GT)
                        || rangeClauseDO.CompareSymbol.Equals(CompareTokens.GTE))
                {
                    if (summaryCoverage.EnumDiscretes.Any())
                    {
                        var capturedEnums = rangeClauseDO.CompareSymbol.Equals(CompareTokens.GT) ?
                            summaryCoverage.EnumDiscretes.Where(en => en > rangeClauseDO.SingleValue)
                            : summaryCoverage.EnumDiscretes.Where(en => en >= rangeClauseDO.SingleValue);

                        isUnreachable = capturedEnums.All(en => summaryCoverage.SingleValues.Contains(en));
                    }
                    else
                    {
                        isUnreachable = summaryCoverage.IsGTMin != null && summaryCoverage.IsGTMin <= rangeClauseDO.SingleValue;
                    }
                }
                else if (CompareTokens.EQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    isUnreachable = SingleValueIsHandledPreviously(rangeClauseDO.SingleValue, summaryCoverage);
                }
                else if (CompareTokens.NEQ.Equals(rangeClauseDO.CompareSymbol))
                {
                    if (rangeClauseDO.TypeNameTarget.Equals(Tokens.Boolean))
                    {
                        isUnreachable = (rangeClauseDO.SingleValue == VBAValue.False ?
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

        private bool SingleValueIsHandledPreviously(VBAValue theValue, SummaryCaseCoverage summaryClauses)
        {
            if (theValue.OriginTypeName.Equals(Tokens.Boolean))
            {
                return summaryClauses.SingleValues.Any(val => val.AsBoolean() == theValue.AsBoolean());
            }
            else
            {
                return summaryClauses.IsLTMax != null && theValue < summaryClauses.IsLTMax
                    || summaryClauses.IsGTMin != null && theValue > summaryClauses.IsGTMin
                    || summaryClauses.SingleValues.Contains(theValue)
                    || summaryClauses.Ranges.Where(rg => theValue.IsWithin(rg.Item1, rg.Item2)).Any()
                    || (summaryClauses.EnumDiscretes.Any() && !summaryClauses.EnumDiscretes.Contains(theValue));
            }
        }

        private SummaryCaseCoverage UpdateSummaryIsClauseLimits(VBAValue theValue, string compareSymbol, SummaryCaseCoverage priorHandlers)
        {
            var isIntegerType = IsIntegerNumberType(theValue.UseageTypeName);
            var isBooleanType = theValue.UseageTypeName.Equals(Tokens.Boolean);

            if (compareSymbol.Equals(CompareTokens.LT) || compareSymbol.Equals(CompareTokens.LTE))
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
                if (isIntegerType)
                {
                    //For Integer Numbers x >= 6 is the same as x > 5
                    //and leads to simpler compares with ValueRanges
                    if(CompareTokens.GTE == compareSymbol)
                    {
                        priorHandlers.IsGTMin = theValue - VBAValue.Unity;
                    }
                    else
                    {
                        priorHandlers.IsLTMax = theValue + VBAValue.Unity;
                    }
                }
                else if (!priorHandlers.SingleValues.Contains(theValue))
                {
                    priorHandlers.SingleValues.Add(theValue);
                }
            }
            return priorHandlers;
        }

        private SummaryCaseCoverage UpdateSummaryDataRanges(SummaryCaseCoverage summaryCoverage, RangeClauseDataObject rangeClauseDO)
        {
            if (!rangeClauseDO.CanBeInspected || !rangeClauseDO.IsValueRange) { return summaryCoverage; }

            if (rangeClauseDO.TypeNameTarget.Equals(Tokens.Boolean))
            {
                if (rangeClauseDO.MinValue != VBAValue.Zero || rangeClauseDO.MaxValue != VBAValue.Zero)
                {
                    summaryCoverage.SingleValues.Add(VBAValue.True);
                }
                if (VBAValue.Zero.IsWithin(rangeClauseDO.MinValue, rangeClauseDO.MaxValue))
                {
                    summaryCoverage.SingleValues.Add(VBAValue.False);
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

        private SummaryCaseCoverage UpdateSummaryDataSingleValues(SummaryCaseCoverage summaryClauses, RangeClauseDataObject rangeClauseDO)
        {
            if (!rangeClauseDO.CanBeInspected || rangeClauseDO.SingleValue == null) { return summaryClauses; }

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
            return summaryClauses;
        }

        private bool EvaluateCaseElseAccessibility(SummaryCaseCoverage summaryClauses, string typeName)
        {
            if(summaryClauses.CaseElseIsUnreachable) { return summaryClauses.CaseElseIsUnreachable; }

            if (typeName.Equals(Tokens.Boolean))
            {
                return summaryClauses.SingleValues.Any(val => val == VBAValue.Zero) && summaryClauses.SingleValues.Any(val => val != VBAValue.Zero)
                    || summaryClauses.IsLTMax != null && summaryClauses.IsLTMax > VBAValue.False
                    || summaryClauses.IsGTMin != null && summaryClauses.IsGTMin < VBAValue.True
                    || summaryClauses.IsLTMax != null && summaryClauses.IsLTMax == VBAValue.False && summaryClauses.SingleValues.Any(sv => sv == VBAValue.False)
                    || summaryClauses.IsGTMin != null && summaryClauses.IsGTMin == VBAValue.True && summaryClauses.SingleValues.Any(sv => sv == VBAValue.True);
            }

            if (summaryClauses.IsLTMax != null && summaryClauses.IsGTMin != null)
            {
                if (summaryClauses.IsLTMax > summaryClauses.IsGTMin
                        || (summaryClauses.IsLTMax >= summaryClauses.IsGTMin
                                && summaryClauses.SingleValues.Contains(summaryClauses.IsLTMax)))
                {
                    return true;
                }

                else if(summaryClauses.Ranges.Count > 0)
                {
                    if (!IsIntegerNumberType(summaryClauses.IsLTMax.OriginTypeName))
                    {
                        return false;
                    }

                    var remainingValues = new List<long>();
                    for (var idx = summaryClauses.IsLTMax.AsLong().Value; idx < summaryClauses.IsGTMin.AsLong().Value; idx++)
                    {
                        remainingValues.Add(idx);
                    }
                    remainingValues.RemoveAll(rv => summaryClauses.Ranges.Any(rg => rg.Item1.AsLong().Value <= rv && rg.Item2.AsLong().Value >= rv));
                    if (remainingValues.Any())
                    {
                        remainingValues.RemoveAll(rv => summaryClauses.SingleValues.Contains(new VBAValue(rv, Tokens.Long)));
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

        private SelectStmtDataObject UpdateTypeNames(SelectStmtDataObject selectStmtDO, string typeName, string baseTypeName = "")
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

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
        #endregion
    }
}