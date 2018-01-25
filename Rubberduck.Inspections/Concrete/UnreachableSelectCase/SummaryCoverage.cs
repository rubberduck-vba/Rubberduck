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
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    internal static class CompareTokens
    {
        public static readonly string EQ = "=";
        public static readonly string NEQ = "<>";
        public static readonly string LT = "<";
        public static readonly string LTE = "<=";
        public static readonly string GT = ">";
        public static readonly string GTE = ">=";
    }

    internal static class MathTokens
    {
        public static readonly string MULT = "*";
        public static readonly string DIV = "/";
        public static readonly string ADD = "+";
        public static readonly string SUBTRACT = "-";
        public static readonly string POW = "^";
        public static readonly string MOD = Tokens.Mod;
    }

    public struct SummaryCaseCoverage
    {
        public UnreachableCaseInspectionValue IsLT;
        public UnreachableCaseInspectionValue IsGT;
        public HashSet<UnreachableCaseInspectionValue> SingleValues;
        public List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>> Ranges;
        public bool CaseElseIsUnreachable;
        public List<string> RangeClausesAsText;
    }

    public class SelectContext : ISelectStmtClause
    {
        private ParserRuleContext _context;
        private List<IParseTree> _contextsOfInterest;

        public SelectContext(ParserRuleContext context)
        {
            _context = context;
            _contextsOfInterest = _context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
        }

        public List<IParseTree> ContextsOfInterest => _contextsOfInterest;

        public void Accept(ISelectStmtClauseVisitor visitor)
        {
            if (_context is VBAParser.LExprContext)
            {
                visitor.Visit((VBAParser.LExprContext)_context);
            }
            else if (_context is VBAParser.LiteralExprContext)
            {
                visitor.Visit((VBAParser.LiteralExprContext)_context);
            }
            else if (_context is VBAParser.SelectCaseStmtContext)
            {
                visitor.Visit((VBAParser.SelectCaseStmtContext)_context);
            }
            else if (_context is VBAParser.SelectExpressionContext)
            {
                visitor.Visit((VBAParser.SelectExpressionContext)_context);
            }
            else if (_context is VBAParser.CaseClauseContext)
            {
                visitor.Visit((VBAParser.CaseClauseContext)_context);
            }
            else if (_context is VBAParser.RangeClauseContext)
            {
                visitor.Visit((VBAParser.RangeClauseContext)_context);
            }
            else if (_context is VBAParser.SelectStartValueContext)
            {
                visitor.Visit((VBAParser.SelectStartValueContext)_context);
            }
            else if (_context is VBAParser.SelectEndValueContext)
            {
                visitor.Visit((VBAParser.SelectEndValueContext)_context);
            }
            else if (_context is VBAParser.RelationalOpContext)
            {
                visitor.Visit((VBAParser.RelationalOpContext)_context);
            }
            else if (_context is VBAParser.MultOpContext)
            {
                visitor.Visit((VBAParser.MultOpContext)_context);
            }
            else if (_context is VBAParser.AddOpContext)
            {
                visitor.Visit((VBAParser.AddOpContext)_context);
            }
            else if (_context is VBAParser.PowOpContext)
            {
                visitor.Visit((VBAParser.PowOpContext)_context);
            }
            else if (_context is VBAParser.ModOpContext)
            {
                visitor.Visit((VBAParser.ModOpContext)_context);
            }
            else if (_context is VBAParser.UnaryMinusOpContext)
            {
                visitor.Visit((VBAParser.UnaryMinusOpContext)_context);
            }
            else if (_context is VBAParser.LogicalAndOpContext)
            {
                visitor.Visit((VBAParser.LogicalAndOpContext)_context);
            }
            else if (_context is VBAParser.LogicalOrOpContext)
            {
                visitor.Visit((VBAParser.LogicalOrOpContext)_context);
            }
            else if (_context is VBAParser.LogicalXorOpContext)
            {
                visitor.Visit((VBAParser.LogicalXorOpContext)_context);
            }
            else if (_context is VBAParser.LogicalEqvOpContext)
            {
                visitor.Visit((VBAParser.LogicalEqvOpContext)_context);
            }
            //else if (_context is VBAParser.LogicalImpOpContext)
            //{
            //    visitor.Visit((VBAParser.LogicalImpOpContext)_context);
            //}
            else if (_context is VBAParser.LogicalNotOpContext)
            {
                visitor.Visit((VBAParser.LogicalNotOpContext)_context);
            }
            else if (_context is VBAParser.ParenthesizedExprContext)
            {
                visitor.Visit((VBAParser.ParenthesizedExprContext)_context);
            }
        }

        public ParserRuleContext Context => _context;

        public IEnumerable<TContext> GetDescendents<TContext>() where TContext : ParserRuleContext
        {
            return Context.GetDescendents<TContext>();
        }

        public IEnumerable<IParseTree> Children()
        {
            return Context.children;
        }
    }

    public class SummaryCoverage : ISelectStmtClauseVisitor
    {
        private SummaryCaseCoverage _summaryCaseCoverage;
        private Dictionary<ParserRuleContext, UnreachableCaseInspectionValue> _constantContexts;
        private Dictionary<ParserRuleContext, UnreachableCaseInspectionValue> _variableContexts;
        private RubberduckParserState _state;
        public SummaryCoverage(RubberduckParserState state)
        {
            _state = state;
            _constantContexts = new Dictionary<ParserRuleContext, UnreachableCaseInspectionValue>(); ;
            _variableContexts = new Dictionary<ParserRuleContext, UnreachableCaseInspectionValue>(); ;
            _summaryCaseCoverage = new SummaryCaseCoverage
            {
                IsGT = null,
                IsLT = null,
                SingleValues = new HashSet<UnreachableCaseInspectionValue>(),
                Ranges = new List<Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>(),
                RangeClausesAsText = new List<string>(),
            };
        }

        public SummaryCaseCoverage Summary => _summaryCaseCoverage;
        public string EvaluationTypeName { set; get; } = string.Empty;
        public RubberduckParserState State => _state;
        public Dictionary<ParserRuleContext, UnreachableCaseInspectionValue> ConstantCtxts => _constantContexts;
        public Dictionary<ParserRuleContext, UnreachableCaseInspectionValue> VariableCtxts => _variableContexts;

        //Used to modify logic operators to convert LHS and RHS for expressions like '5 > x' (= 'x < 5')
        private static Dictionary<string, string> AlgebraicLogicalInversions = new Dictionary<string, string>()
        {
            [CompareTokens.EQ] = CompareTokens.EQ,
            [CompareTokens.NEQ] = CompareTokens.NEQ,
            [CompareTokens.LT] = CompareTokens.GT,
            [CompareTokens.LTE] = CompareTokens.GTE,
            [CompareTokens.GT] = CompareTokens.LT,
            [CompareTokens.GTE] = CompareTokens.LTE
        };

        private static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>
            BinaryMathOps = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
            {
                [MathTokens.ADD] = (LHS, RHS) => LHS + RHS,
                [MathTokens.SUBTRACT] = (LHS, RHS) => LHS - RHS,
                [MathTokens.MULT] = (LHS, RHS) => LHS * RHS,
                [MathTokens.DIV] = (LHS, RHS) => LHS / RHS,
                [MathTokens.POW] = (LHS, RHS) => UnreachableCaseInspectionValue.Pow(LHS,RHS),
                [MathTokens.MOD] = (LHS, RHS) => LHS % RHS
            };

        private static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>
            BinaryLogicalOps = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
            {
                [CompareTokens.GT] = (LHS, RHS) => LHS > RHS ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [CompareTokens.GTE] = (LHS, RHS) => LHS >= RHS ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [CompareTokens.LT] = (LHS, RHS) => LHS < RHS ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [CompareTokens.LTE] = (LHS, RHS) => LHS <= RHS ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [Tokens.And] = (LHS, RHS) => LHS.AsBoolean().Value && RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [Tokens.Or] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [Tokens.XOr] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
                [Tokens.Not] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False
                //["Eqv"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False
                //["Imp"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
            };

        private static Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>
            UnaryLogicalOps = new Dictionary<string, Func<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>>()
            {
                [Tokens.Not] = (LHS) => !(LHS.AsBoolean().Value) ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False
            };

        private void StoreVisitResult(ParserRuleContext context, UnreachableCaseInspectionValue result)
        {
            if (result.HasValue)
            {
                _constantContexts.Add(context, result);
            }
            else
            {
                _variableContexts.Add(context, result);
            }
        }

        private bool ContextHasResult(ParserRuleContext context)
        {
            return _constantContexts.Keys.Contains(context) || _variableContexts.Keys.Contains(context);
        }

        private void VisitImpl<T>(T context) where T: ParserRuleContext
        {
            if (ContextHasResult(context))
            {
                return;
            }

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            foreach (var ctxt in contextsOfInterest)
            {
                if (ctxt is ParserRuleContext)
                {
                    var selectContext = new SelectContext((ParserRuleContext)ctxt);
                    selectContext.Accept(this);
                }
            }
        }

        public void Visit(VBAParser.SelectCaseStmtContext selectStmt)
        {
            VisitImpl(selectStmt);
        }

        public void Visit(VBAParser.SelectExpressionContext selectStmt)
        {
            VisitImpl(selectStmt);
        }

        public void Visit(VBAParser.CaseClauseContext caseClause)
        {

        }

        public void Visit(VBAParser.RangeClauseContext rangeClause)
        {
            VisitImpl(rangeClause);

            //Range of values 35 To 70
            if (rangeClause.HasChildToken(Tokens.To))
            {
                var startContext = rangeClause.GetChild<VBAParser.SelectStartValueContext>();
                var endContext = rangeClause.GetChild<VBAParser.SelectEndValueContext>();
                if (_constantContexts.TryGetValue(startContext, out UnreachableCaseInspectionValue startVal) &&
                 _constantContexts.TryGetValue(endContext, out UnreachableCaseInspectionValue endVal))
                {
                    if (startVal <= endVal)
                    {
                        Summary.Ranges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(startVal, endVal));
                    }
                    else
                    {
                        Summary.Ranges.Add(new Tuple<UnreachableCaseInspectionValue, UnreachableCaseInspectionValue>(endVal, startVal));
                    }
                }
            }
            else //single value
            {
                var ctxts = rangeClause.children.Where(ch => ch is ParserRuleContext
                                && _constantContexts.Keys.Contains((ParserRuleContext)ch)
                                && _constantContexts[(ParserRuleContext)ch].HasValue);

                if (ctxts.Any() && ctxts.Count() == 1 && rangeClause.HasChildToken(Tokens.Is) )
                {
                    var compOpContext = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>();
                    AddIsClauseResult(compOpContext.GetText(), _constantContexts[(ParserRuleContext)ctxts.First()]);
                }
                else if(rangeClause.children.Any() && rangeClause.children.Count() == 1 && rangeClause.children.First() is VBAParser.RelationalOpContext)
                {
                    var relOpCtxt = (ParserRuleContext)rangeClause.children.First();
                    if(!_constantContexts.Keys.Contains(relOpCtxt))
                    {
                        var relOpContexts = relOpCtxt.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
                        UnreachableCaseInspectionValue LHS = null;
                        UnreachableCaseInspectionValue RHS = null;
                        for (var idx = 0; idx < relOpContexts.Count(); idx++)
                        {
                            var ctxt = relOpContexts[idx];
                            if (_constantContexts.Keys.Contains(ctxt) || _variableContexts.Keys.Contains(ctxt))
                            {
                                if (LHS is null)
                                {
                                    LHS = _constantContexts.Keys.Contains(ctxt) ?
                                                _constantContexts[(ParserRuleContext)ctxt]
                                                : _variableContexts[(ParserRuleContext)ctxt];
                                }
                                else if (RHS is null)
                                {
                                    RHS = _constantContexts.Keys.Contains(ctxt) ?
                                                _constantContexts[(ParserRuleContext)ctxt]
                                                : _variableContexts[(ParserRuleContext)ctxt];
                                }
                            }
                        }
                        if (LHS != null && RHS != null)
                        {
                            var opSymbol = relOpCtxt.children.Where(ch => BinaryLogicalOps.Keys.Contains(ch.GetText())).First().GetText();
                            if (BinaryLogicalOps.ContainsKey(opSymbol))
                            {
                                if (LHS.HasValue)
                                {
                                    opSymbol = AlgebraicLogicalInversions[opSymbol];
                                }
                                var result = LHS.HasValue ? LHS : RHS;

                                if (result.HasValue)
                                {
                                    AddIsClauseResult(opSymbol, result);
                                }
                            }
                        }
                    }
                }
                else if (ctxts.Any() && ctxts.Count() == 1)
                {
                    AddSingleValue(_constantContexts[(ParserRuleContext)ctxts.First()]);
                }
            }
        }

        public void Visit(VBAParser.LExprContext context)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            UnreachableCaseInspectionValue result = null;
            var lexprTypeName = this.EvaluationTypeName;
            if (TryGetTheLExprValue(context, out string lexprValue, ref lexprTypeName))
            {
                result = this.EvaluationTypeName.Length > 0 ? new UnreachableCaseInspectionValue(lexprValue, this.EvaluationTypeName) : new UnreachableCaseInspectionValue(lexprValue, lexprTypeName);
            }
            else
            {
                var smplName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference idRef))
                {
                    var theTypeName = GetBaseTypeForDeclaration(idRef.Declaration);
                    result = new UnreachableCaseInspectionValue(context.GetText(), theTypeName);
                }
            }

            if (result != null)
            {
                StoreVisitResult(context, result);
            }
        }

        public void Visit(VBAParser.LiteralExprContext context)
        {
            if (!ContextHasResult(context))
            {
                var result = new UnreachableCaseInspectionValue(context.GetText(), this.EvaluationTypeName);
                StoreVisitResult(context, result);
            }
        }

        private void AddIsClauseResult(string compareOperator, UnreachableCaseInspectionValue result)
        {
            if (compareOperator.Equals(CompareTokens.LT))
            {
                AddIsLT(result);
            }
            else if (compareOperator.Equals(CompareTokens.LTE))
            {
                AddIsLT(result);
                AddSingleValue(result);
            }
            else if (compareOperator.Equals(CompareTokens.GT))
            {
                AddIsGT(result);
            }
            else if (compareOperator.Equals(CompareTokens.GTE))
            {
                AddIsGT(result);
                AddSingleValue(result);
            }
            else if (compareOperator.Equals(CompareTokens.EQ))
            {
                AddSingleValue(result);
            }
            else if (compareOperator.Equals(CompareTokens.NEQ))
            {
                AddIsLT(result);
                AddIsGT(result);
            }
            else
            {
                Debug.Assert(false, "Unrecognized comparison symbol for Is Clause");
            }
        }

        public void Visit(VBAParser.MultOpContext context)
        {
            VisitMathOpBinary(context);
        }

        public void Visit(VBAParser.AddOpContext context)
        {
            VisitMathOpBinary(context);
        }

        public void Visit(VBAParser.PowOpContext context)
        {
            VisitMathOpBinary(context);
        }

        public void Visit(VBAParser.ModOpContext context)
        {
            VisitMathOpBinary(context);
        }

        public void Visit(VBAParser.UnaryMinusOpContext context)
        {
            VisitImpl(context);

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_constantContexts.Keys.Contains(ctxt) && _constantContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    StoreVisitResult(context, _constantContexts[(ParserRuleContext)ctxt].AdditiveInverse);
                }
            }
        }

        public void VisitMathOpBinary(ParserRuleContext context)
        {
            VisitImpl(context);

            UnreachableCaseInspectionValue LHS = null;
            UnreachableCaseInspectionValue RHS = null;
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_constantContexts.Keys.Contains(ctxt) && _constantContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    if (LHS is null)
                    {
                        LHS = _constantContexts[(ParserRuleContext)ctxt];
                    }
                    else if (RHS is null)
                    {
                        RHS = _constantContexts[(ParserRuleContext)ctxt];
                    }
                }
            }

            if (LHS != null && LHS.HasValue && RHS != null && RHS.HasValue)
            {
                var opSymbol = context.children.Where(ch => BinaryMathOps.Keys.Contains(ch.GetText())).First().GetText();
                if (BinaryMathOps.ContainsKey(opSymbol))
                {
                    StoreVisitResult(context, BinaryMathOps[opSymbol](LHS, RHS));
                }
            }
        }

        public void Visit(VBAParser.LogicalAndOpContext context)
        {
            VisitLogicalOperation(context);
        }

        public void Visit(VBAParser.LogicalOrOpContext context)
        {
            VisitLogicalOperation(context);
        }

        public void Visit(VBAParser.RelationalOpContext context)
        {
            VisitLogicalOperation(context);
        }

        public void Visit(VBAParser.LogicalXorOpContext context)
        {
            VisitLogicalOperation(context);
        }

        public void Visit(VBAParser.LogicalEqvOpContext context)
        {
            VisitLogicalOperation(context);
        }

        //public void Visit(VBAParser.LogicalImpOpContext context)
        //{
        //    //VisitLogicalOpBinary(context);
        //}

        public void Visit(VBAParser.LogicalNotOpContext context)
        {
            VisitLogicalOperation(context, false);
        }

        public void VisitLogicalOperation(ParserRuleContext context, bool isBinaryOp = true)
        {
            VisitImpl(context);

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            UnreachableCaseInspectionValue LHS = null;
            UnreachableCaseInspectionValue RHS = null;
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_constantContexts.Keys.Contains(ctxt) && _constantContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    if (LHS is null)
                    {
                        LHS = _constantContexts[(ParserRuleContext)ctxt];
                    }
                    else if (RHS is null && isBinaryOp)
                    {
                        RHS = _constantContexts[(ParserRuleContext)ctxt];
                    }
                }
            }

            if (isBinaryOp)
            {
                if (LHS != null && LHS.HasValue && RHS != null && RHS.HasValue)
                {
                    var opSymbol = context.children.Where(ch => BinaryLogicalOps.Keys.Contains(ch.GetText())).First().GetText();
                    if (BinaryLogicalOps.ContainsKey(opSymbol))
                    {
                        var result = new UnreachableCaseInspectionValue(BinaryLogicalOps[opSymbol](LHS, RHS).ToString(), EvaluationTypeName);
                        StoreVisitResult(context, result);
                    }
                }
            }
            else
            {
                if (LHS != null && LHS.HasValue)
                {
                    var opSymbol = context.children.Where(ch => BinaryLogicalOps.Keys.Contains(ch.GetText())).First().GetText();
                    if (UnaryLogicalOps.ContainsKey(opSymbol))
                    {
                        var result = new UnreachableCaseInspectionValue(UnaryLogicalOps[opSymbol](LHS).ToString(), EvaluationTypeName);
                        StoreVisitResult(context, result);
                    }
                }
            }
        }

        public void Visit(VBAParser.ParenthesizedExprContext context)
        {
            VisitImpl(context);

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_constantContexts.Keys.Contains(ctxt) && _constantContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    StoreVisitResult(context, _constantContexts[(ParserRuleContext)ctxt]);
                }
            }
        }

        public void Visit(VBAParser.SelectEndValueContext context)
        {
            VisitImpl(context);
            StartEndContextResult(context);
        }

        public void Visit(VBAParser.SelectStartValueContext context)
        {
            VisitImpl(context);
            StartEndContextResult(context);
        }

        private void StartEndContextResult(ParserRuleContext context)
        {
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            foreach (var ctxt in contextsOfInterest)
            {
                if (_constantContexts.Keys.Contains(ctxt))
                {
                    if (!_constantContexts.Keys.Contains(context))
                    {
                        StoreVisitResult(context, _constantContexts[(ParserRuleContext)ctxt]);
                    }
                }
            }
        }

        private bool TryGetTheLExprValue(VBAParser.LExprContext ctxt, out string expressionValue, ref string typeName)
        {
            expressionValue = string.Empty;
            if (ctxt.TryGetChildContext(out VBAParser.MemberAccessExprContext memberAccess))
            {
                var member = memberAccess.GetChild<VBAParser.UnrestrictedIdentifierContext>();

                if (TryGetIdentifierReferenceForContext(member, out IdentifierReference idRef))
                {
                    var dec = idRef.Declaration;
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
                //var memberDeclarations = State.DeclarationFinder.AllUserDeclarations.Where(dec => dec.IdentifierName.Equals(member.GetText()));

                //foreach (var dec in memberDeclarations)
                //{
                //    if (dec.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                //    {
                //        var theCtxt = dec.Context;
                //        if (theCtxt is EnumerationStmt_ConstantContext)
                //        {
                //            expressionValue = GetConstantDeclarationValue(dec);
                //            typeName = dec.AsTypeIsBaseType ? dec.AsTypeName : dec.AsTypeDeclaration.AsTypeName;
                //            return true;
                //        }
                //    }
                //}
                return false;
            }

            if (ctxt.TryGetChildContext(out VBAParser.SimpleNameExprContext smplName))
            {
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference rangeClauseIdentifierReference))
                {
                    var declaration = rangeClauseIdentifierReference.Declaration;
                    if (declaration.DeclarationType.HasFlag(DeclarationType.Constant)
                        || declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                    {
                        expressionValue = GetConstantDeclarationValue(declaration);
                        typeName = declaration.AsTypeName;
                        return true;
                    }
                }
                //var identifierReferences = (State.DeclarationFinder.MatchName(smplName.GetText()).Select(dec => dec.References)).SelectMany(rf => rf);
                //var rangeClauseReferences = identifierReferences.Where(rf => rf.Context.Parent == smplName);

                //var rangeClauseIdentifierReference = rangeClauseReferences.Any() ? rangeClauseReferences.First() : null;
                //if (rangeClauseIdentifierReference != null)
                //{
                //    var declaration = rangeClauseIdentifierReference.Declaration;
                //    if (declaration.DeclarationType.HasFlag(DeclarationType.Constant)
                //        || declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                //    {
                //        expressionValue = GetConstantDeclarationValue(declaration);
                //        typeName = declaration.AsTypeName;
                //        return true;
                //    }
                //}
            }
            return false;
        }

        private bool TryGetIdentifierReferenceForContext<T>(T context, out IdentifierReference idRef) where T : ParserRuleContext
        {
            idRef = null;
            var identifierReferences = (State.DeclarationFinder.MatchName(context.GetText()).Select(dec => dec.References)).SelectMany(rf => rf);
            if (identifierReferences.Any())
            {
                idRef = identifierReferences.First(rf => rf.Context == context);
                return true;
            }
            return false;
        }

        private string GetConstantDeclarationValue(Declaration valueDeclaration)
        {
            var contextsOfInterest = GetRHSContexts(valueDeclaration.Context.children.ToList());
            foreach (var child in contextsOfInterest)
            {
                //if (IsMathOperation(child))
                //{
                //    var parentData = new Dictionary<IParseTree, ExpressionEvaluationDataObject>();
                //    var exprEval = new ExpressionEvaluationDataObject
                //    {
                //        IsUnaryOperation = IsUnaryMathOperation(child),
                //        Operator = CompareTokens.EQ,
                //        CanBeInspected = true,
                //        TypeNameTarget = valueDeclaration.AsTypeName,
                //        SelectCaseRefName = valueDeclaration.IdentifierName
                //    };

                //    parentData = AddEvaluationData(parentData, child, exprEval);
                //    return ResolveContextValue(parentData, child).First().Value.Result.AsString();
                //}

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

        private static List<ParserRuleContext> GetRHSContexts(List<IParseTree> contexts)
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

        private static string GetBaseTypeForDeclaration(Declaration declaration)
        {
            if (!declaration.AsTypeIsBaseType)
            {
                return GetBaseTypeForDeclaration(declaration.AsTypeDeclaration);
            }
            return declaration.AsTypeName;
        }

        private void AddIsLT(UnreachableCaseInspectionValue isLT)
        {
            if(this.Summary.IsLT == null)
            {
                _summaryCaseCoverage.IsLT = isLT;
            }
            else if(_summaryCaseCoverage.IsLT < isLT)
            {
                _summaryCaseCoverage.IsLT = isLT;
            }
        }

        private void AddIsGT(UnreachableCaseInspectionValue isGT)
        {
            if (Summary.IsGT == null)
            {
                _summaryCaseCoverage.IsGT = isGT;
            }
            else if (_summaryCaseCoverage.IsGT > isGT)
            {
                _summaryCaseCoverage.IsGT = isGT;
            }
        }

        private void AddSingleValue(UnreachableCaseInspectionValue singleValue)
        {
            this.Summary.SingleValues.Add(singleValue);
        }

        private void AddRange(HashSet<UnreachableCaseInspectionValue> singleValues)
        {

        }

        private void AddRangeClauseText(string rangeClausesAsText)
        {

        }

        private void AddRange(IEnumerable<string> rangeClausesAsText)
        {

        }

        private void AddValueRange(List<Tuple<UnreachableCaseInspectionValue,UnreachableCaseInspectionValue>> valueRange)
        {

        }
    }
}
