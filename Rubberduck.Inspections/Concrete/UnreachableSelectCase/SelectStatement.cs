using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{

    public static class CompareTokens
    {
        public static readonly string EQ = "=";
        public static readonly string NEQ = "<>";
        public static readonly string LT = "<";
        public static readonly string LTE = "<=";
        public static readonly string GT = ">";
        public static readonly string GTE = ">=";
    }

    public struct ExpressionEvaluationDataObject
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

    public enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse, NoResult };

    public struct CaseClauseDataObject
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

    public struct RangeClauseDataObject
    {
        public VBAParser.RangeClauseContext Context;
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

        public RangeClauseDataObject(VBAParser.RangeClauseContext ctxt, string targetTypeName)
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

    public struct SelectStmtDataObject
    {
        public VBAParser.SelectCaseStmtContext SelectStmtContext;
        public VBAParser.SelectExpressionContext SelectExpressionContext;
        public string BaseTypeName;
        public string AsTypeName;
        public string IdReferenceName;
        public List<CaseClauseDataObject> CaseClauseDOs;
        public VBAParser.CaseElseClauseContext CaseElseContext;
        public SummaryCaseCoverage SummaryCaseClauses;
        public bool CanBeInspected;

        public SelectStmtDataObject(QualifiedContext<ParserRuleContext> selectStmtCtxt)
        {
            SelectStmtContext = (VBAParser.SelectCaseStmtContext)selectStmtCtxt.Context;
            IdReferenceName = string.Empty;
            BaseTypeName = Tokens.Variant;
            AsTypeName = Tokens.Variant;
            CaseClauseDOs = new List<CaseClauseDataObject>();
            CaseElseContext = SelectStmtContext.FindChild<VBAParser.CaseElseClauseContext>();
            SummaryCaseClauses = new SummaryCaseCoverage
            {
                IsGT = null,
                IsLT = null,
                SingleValues = new HashSet<UnreachableCaseInspectionValue>(),
                Ranges = new List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>(),
                RangeClausesAsText = new List<string>(),
            };
            CanBeInspected = SelectStatement.TryGetChildContext(SelectStmtContext, out SelectExpressionContext);
        }
    }

    public class SelectStatement
    {
        private SummaryCoverage _summaryCoverage;
        private readonly RubberduckParserState _state;
        private RubberduckParserState State => _state;
        private SelectStmtDataObject _selectStmtDO;

        public string EvaluationType => _selectStmtDO.BaseTypeName;


        public SelectStatement(RubberduckParserState state, QualifiedContext<ParserRuleContext> selectStmtCtxt)
        {
            _summaryCoverage = new SummaryCoverage();
            _state = state;
            _selectStmtDO = new SelectStmtDataObject(selectStmtCtxt);

            _selectStmtDO = ResolveSelectStmtInspectionType(_selectStmtDO);

            if (_selectStmtDO.CanBeInspected)
            {
                _selectStmtDO.CaseClauseDOs = _selectStmtDO.SelectStmtContext.GetChildren<VBAParser.CaseClauseContext>()
                    .Select(cc => CreateCaseClauseDataObject(cc, _selectStmtDO.BaseTypeName)).ToList();
            }
        }

        private CaseClauseDataObject CreateCaseClauseDataObject(VBAParser.CaseClauseContext ctxt, string targetTypeName)
        {
            var caseClauseDO = new CaseClauseDataObject(ctxt);
            var rangeClauseContexts = ctxt.GetChildren<VBAParser.RangeClauseContext>();
            foreach (var rangeClauseCtxt in rangeClauseContexts)
            {
                var rgC = new RangeClauseDataObject(rangeClauseCtxt, targetTypeName);
                caseClauseDO.RangeClauseDOs.Add(rgC);
            }
            return caseClauseDO;
        }

        private SelectStmtDataObject ResolveSelectStmtInspectionType(SelectStmtDataObject selectStmtDO)
        {
                if (!_selectStmtDO.CanBeInspected)
                {
                    return selectStmtDO;
                }
                //if (!ContextCanBeEvaluated(selectStmtDO.SelectExpressionContext, selectStmtDO.IdReferenceName))
                //{
                //    return InferTheSelectStmtType(selectStmtDO);
                //}

                if (selectStmtDO.SelectExpressionContext.GetDescendents()
                .Any(desc => IsBinaryLogicalOperation(desc) || IsUnaryLogicalOperator(desc)))
            {
                return SetTheTypeNames(selectStmtDO, Tokens.Boolean);
            }

            var firstLExpr = selectStmtDO.SelectExpressionContext.GetDescendent<VBAParser.LExprContext>();
            if (firstLExpr == null)
            {
                return InferTheSelectStmtType(selectStmtDO);
            }

            var expression = firstLExpr.GetDescendent<VBAParser.SimpleNameExprContext>().GetText();
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

        private bool ContextCanBeEvaluated(ParserRuleContext context, string refName)
        {
            var canBeInspected = true;
            //var descs = context.GetDescendents();
            var ops = context.GetDescendents().Where(desc => (desc is ParserRuleContext) && (IsBinaryMathOperation(desc) || IsUnaryMathOperation(desc)));
            foreach (var op in ops)
            {
                var lExpressions = ((ParserRuleContext)op).FindChildren<VBAParser.LExprContext>();
                var mathOnTheSelectCaseVariable = lExpressions.Any(lex => lex.GetText().Equals(refName));
                var mathOnNonConstants = lExpressions.Any(lex => !(CreateValue(lex, Tokens.Variant).HasValue));

                if (mathOnTheSelectCaseVariable || mathOnNonConstants)
                {
                    canBeInspected = false;
                }
            }
            return canBeInspected;
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

        public static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> MathOperations = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
        {
            ["*"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS * RHS; },
            ["/"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS / RHS; },
            ["+"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS + RHS; },
            ["-"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS - RHS; },
            ["^"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS ^ RHS; },
            ["Mod"] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return LHS % RHS; }
        };

        public static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> CompareOperations = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
        {
            [CompareTokens.EQ] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return new UnreachableCaseInspectionValue(LHS == RHS ? Tokens.True : Tokens.False, Tokens.Boolean); },
            [CompareTokens.NEQ] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return new UnreachableCaseInspectionValue(LHS != RHS ? Tokens.True : Tokens.False, Tokens.Boolean); },
            [CompareTokens.LT] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return new UnreachableCaseInspectionValue(LHS < RHS ? Tokens.True : Tokens.False, Tokens.Boolean); },
            [CompareTokens.LTE] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return new UnreachableCaseInspectionValue(LHS <= RHS ? Tokens.True : Tokens.False, Tokens.Boolean); },
            [CompareTokens.GT] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return new UnreachableCaseInspectionValue(LHS > RHS ? Tokens.True : Tokens.False, Tokens.Boolean); },
            [CompareTokens.GTE] = delegate (UnreachableCaseInspectionValue LHS, UnreachableCaseInspectionValue RHS) { return new UnreachableCaseInspectionValue(LHS >= RHS ? Tokens.True : Tokens.False, Tokens.Boolean); }
        };

        public static bool TryGetChildContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : ParserRuleContext
        {
            opCtxt = ctxt.FindChild<U>();
            return opCtxt != null;
        }

        private bool TryInferTypeFromRangeClauseContent(SelectStmtDataObject selectStmtDO, out string typeName)
        {
            typeName = selectStmtDO.BaseTypeName;

            if (!selectStmtDO.BaseTypeName.Equals(Tokens.Variant)) { return false; }

            var rangeCtxts = selectStmtDO.SelectStmtContext.GetChildren<VBAParser.CaseClauseContext>()
                .Select(cc => cc.GetChildren<VBAParser.RangeClauseContext>())
                .SelectMany(rgCtxt => rgCtxt);

            var typeNames = rangeCtxts.SelectMany(rgCtxt => rgCtxt.GetDescendents())
                .Where(desc => desc is VBAParser.LiteralExprContext || desc is VBAParser.LExprContext)
                    .Select(exprCtxt => EvaluateContextTypeName((VBAParser.ExpressionContext)exprCtxt, selectStmtDO))
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

        private string EvaluateContextTypeName(VBAParser.ExpressionContext ctxt, SelectStmtDataObject selectStmtDO)
        {
            var val = CreateValue(ctxt, selectStmtDO.BaseTypeName);
            return val.HasValue ? selectStmtDO.BaseTypeName : val.DerivedTypeName;
        }

        private UnreachableCaseInspectionValue CreateValue(VBAParser.ExpressionContext ctxt, string typeName = "")
        {
            if (ctxt is VBAParser.LExprContext)
            {
                var lexprTypeName = typeName;
                if (TryGetTheLExprValue((VBAParser.LExprContext)ctxt, out string lexprValue, ref lexprTypeName))
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
            else if (ctxt is VBAParser.LiteralExprContext)
            {
                return new UnreachableCaseInspectionValue(ctxt.GetText(), typeName);
            }
            return null;
        }

        public bool TryGetTheLExprValue(VBAParser.LExprContext ctxt, out string expressionValue, ref string typeName)
        {
            expressionValue = string.Empty;
            if (SelectStatement.TryGetChildContext(ctxt, out VBAParser.MemberAccessExprContext member))
            {
                var smplNameMemberRHS = member.FindChild<VBAParser.UnrestrictedIdentifierContext>();
                var memberDeclarations = State.DeclarationFinder.AllUserDeclarations.Where(dec => dec.IdentifierName.Equals(smplNameMemberRHS.GetText()));

                foreach (var dec in memberDeclarations)
                {
                    if (dec.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        var theCtxt = dec.Context;
                        if (theCtxt is VBAParser.EnumerationStmt_ConstantContext)
                        {
                            expressionValue = GetConstantDeclarationValue(dec);
                            typeName = dec.AsTypeIsBaseType ? dec.AsTypeName : dec.AsTypeDeclaration.AsTypeName;
                            return true;
                        }
                    }
                }
                return false;
            }
            else if (SelectStatement.TryGetChildContext(ctxt, out VBAParser.SimpleNameExprContext smplName))
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

                if (child is VBAParser.LiteralExprContext)
                {
                    if (child.Parent is VBAParser.EnumerationStmt_ConstantContext)
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

        private Dictionary<IParseTree, ExpressionEvaluationDataObject> ResolveContextValue(Dictionary<IParseTree, ExpressionEvaluationDataObject> contextEvals, ParserRuleContext parentContext)
        {
            var parentData = GetEvaluationData(parentContext, contextEvals);
            foreach (var child in parentContext.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)))
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
                else if (child is VBAParser.LiteralExprContext || child is VBAParser.LExprContext)
                {
                    childData.IsUnaryOperation = true;
                    childData.LHSValue = CreateValue((VBAParser.ExpressionContext)child, childData.TypeNameTarget);
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

        public static List<ParserRuleContext> GetRHSContexts(List<IParseTree> contexts)
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
                if (!(childCtxt is VBAParser.WhiteSpaceContext))
                {
                    contextsOfInterest.Add((ParserRuleContext)childCtxt);
                }
            }
            return contextsOfInterest;
        }

        public static string GetBaseTypeForDeclaration(Declaration declaration)
        {
            if (!declaration.AsTypeIsBaseType)
            {
                return GetBaseTypeForDeclaration(declaration.AsTypeDeclaration);
            }
            return declaration.AsTypeName;
        }

        private SelectStmtDataObject SetTheTypeNames(SelectStmtDataObject selectStmtDO, string typeName, string baseTypeName = "")
        {
            selectStmtDO.AsTypeName = typeName;
            selectStmtDO.BaseTypeName = baseTypeName.Length == 0 ? typeName : baseTypeName;
            return selectStmtDO;
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
            return child is VBAParser.MultOpContext
                || child is VBAParser.AddOpContext
                || child is VBAParser.PowOpContext
                || child is VBAParser.ModOpContext;
        }

        private bool IsBinaryLogicalOperation<T>(T child)
        {
            return child is VBAParser.RelationalOpContext
                || child is VBAParser.LogicalXorOpContext
                || child is VBAParser.LogicalAndOpContext
                || child is VBAParser.LogicalOrOpContext
                || child is VBAParser.LogicalEqvOpContext
                || child is VBAParser.LogicalNotOpContext;
        }

        private bool IsUnaryLogicalOperator<T>(T child)
        {
            return child is VBAParser.LogicalNotOpContext;
        }

        private bool IsUnaryMathOperation<T>(T child)
        {
            return child is VBAParser.UnaryMinusOpContext;
        }

        private bool IsUnaryOperandContext<T>(T child)
        {
            return IsUnaryLogicalOperator(child)
                   || IsUnaryMathOperation(child)
                   || child is VBAParser.ParenthesizedExprContext;
        }
    }
}
