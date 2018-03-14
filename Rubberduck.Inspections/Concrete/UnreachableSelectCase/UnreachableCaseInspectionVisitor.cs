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

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public class IUnreachableCaseInspectionValueFactory
    {
        public IUnreachableCaseInspectionValue Create(string valueToken)
        {
            return new UnreachableCaseInspectionValue(valueToken);
        }

        public IUnreachableCaseInspectionValue Create(string valueToken, string declaredTypeName)
        {
            return new UnreachableCaseInspectionValue(valueToken, declaredTypeName);
        }
    }

    internal class UnreachableCaseInspectionValueVisitor : IParseTreeVisitor<IUnreachableCaseInspectionValue>
    {
        private Dictionary<ParserRuleContext, IUnreachableCaseInspectionValue> _contextValues;
        private RubberduckParserState _state;
        private IUnreachableCaseInspectionValueFactory _inspValueTreeFactory;

        public UnreachableCaseInspectionValueVisitor(RubberduckParserState state, IUnreachableCaseInspectionValueFactory factory, string evaluationTypeName = "")
        {
            _state = state;
            _contextValues = new Dictionary<ParserRuleContext, IUnreachableCaseInspectionValue>();
            EvaluationTypeName = evaluationTypeName ?? string.Empty;
            _inspValueTreeFactory = factory;
        }

        internal static class MathTokens
        {
            public static readonly string MULT = "*";
            public static readonly string DIV = "/";
            public static readonly string ADD = "+";
            public static readonly string SUBTRACT = "-";
            public static readonly string POW = "^";
            public static readonly string MOD = Tokens.Mod;
            public static readonly string ADDITIVE_INVERSE = "-";
        }

        public static Dictionary<string, UnreachableCaseInspectionBinaryOp> BinaryOps = new Dictionary<string, UnreachableCaseInspectionBinaryOp>()
        {
            [MathTokens.MULT] = new UnreachableCaseInspectionBinaryOp(MathTokens.MULT),
            [MathTokens.DIV] = new UnreachableCaseInspectionBinaryOp(MathTokens.DIV),
            [MathTokens.ADD] = new UnreachableCaseInspectionBinaryOp(MathTokens.ADD),
            [MathTokens.SUBTRACT] = new UnreachableCaseInspectionBinaryOp(MathTokens.SUBTRACT),
            [MathTokens.POW] = new UnreachableCaseInspectionBinaryOp(MathTokens.POW),
            [MathTokens.MOD] = new UnreachableCaseInspectionBinaryOp(MathTokens.MOD),
            [CompareTokens.EQ] = new UnreachableCaseInspectionBinaryOp(CompareTokens.EQ),
            [CompareTokens.NEQ] = new UnreachableCaseInspectionBinaryOp(CompareTokens.NEQ),
            [CompareTokens.LT] = new UnreachableCaseInspectionBinaryOp(CompareTokens.LT),
            [CompareTokens.LTE] = new UnreachableCaseInspectionBinaryOp(CompareTokens.LTE),
            [CompareTokens.GT] = new UnreachableCaseInspectionBinaryOp(CompareTokens.GT),
            [CompareTokens.GTE] = new UnreachableCaseInspectionBinaryOp(CompareTokens.GTE),
            [Tokens.And] = new UnreachableCaseInspectionBinaryOp(Tokens.And),
            [Tokens.Or] = new UnreachableCaseInspectionBinaryOp(Tokens.Or),
            [Tokens.XOr] = new UnreachableCaseInspectionBinaryOp(Tokens.XOr),
        };

        public static Dictionary<string, UnreachableCaseInspectionUnaryOp> UnaryOps = new Dictionary<string, UnreachableCaseInspectionUnaryOp>()
        {
            [Tokens.Not] = new UnreachableCaseInspectionUnaryOp(Tokens.Not),
            [MathTokens.ADDITIVE_INVERSE] = new UnreachableCaseInspectionUnaryOp(MathTokens.ADDITIVE_INVERSE)
        };

        public string EvaluationTypeName { set; get; } = string.Empty;
        private RubberduckParserState State => _state;

        public virtual IUnreachableCaseInspectionValue Visit(IParseTree tree)
        {
            if(tree is ParserRuleContext context)
            {
                Visit(context);
                if (_contextValues.ContainsKey(context))
                {
                    return _contextValues[context];
                }
            }
            return  new UnreachableCaseInspectionValue(double.NaN.ToString(), Tokens.Variant);
        }

        public virtual IUnreachableCaseInspectionValue VisitChildren(IRuleNode node)
        {
            if (node is ParserRuleContext context)
            {
                Visit(context);
                if (_contextValues.ContainsKey(context))
                {
                    return _contextValues[context];
                }
            }
            return new UnreachableCaseInspectionValue(double.NaN.ToString(), Tokens.Variant);
        }

        public virtual IUnreachableCaseInspectionValue VisitTerminal(ITerminalNode node)
        {
            return new UnreachableCaseInspectionValue(double.NaN.ToString(), Tokens.Variant);
        }

        public virtual IUnreachableCaseInspectionValue VisitErrorNode(IErrorNode node)
        {
            return new UnreachableCaseInspectionValue(double.NaN.ToString(), Tokens.Variant);
        }

        public void Visit(ParserRuleContext parserRuleContext)
        {
            if (IsUnaryResultContext(parserRuleContext))
            {
                VisitSummaryContextType(parserRuleContext);
            }
            else if (parserRuleContext is VBAParser.LExprContext lExpr)
            {
                Visit(lExpr);
            }
            else if (parserRuleContext is VBAParser.LiteralExprContext litExpr)
            {
                Visit(litExpr);
            }
            else if (parserRuleContext is VBAParser.SelectCaseStmtContext)
            {
                VisitImpl(parserRuleContext);
            }
            else if (parserRuleContext is VBAParser.RangeClauseContext rangeCtxt)
            {
                VisitSummaryContextType(parserRuleContext);
            }
            else if (IsBinaryMathContext(parserRuleContext) || IsBinaryLogicalContext(parserRuleContext))
            {
                VisitBinaryEvaluationContext(parserRuleContext);
            }
            else if (IsUnaryLogicalContext(parserRuleContext))
            {
                VisitUnaryLogicalContext(parserRuleContext);
            }
            else if (parserRuleContext is VBAParser.UnaryMinusOpContext uMinusOp)
            {
                Visit(uMinusOp);
            }
        }

        private void Visit(VBAParser.LExprContext context)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            IUnreachableCaseInspectionValue newResult = null;
            if (TryGetTheLExprValue(context, out string lexprValue, out string declaredType))
            {
                newResult = _inspValueTreeFactory.Create(lexprValue, declaredType);
            }
            else
            {
                var smplNameExprTypeName = this.EvaluationTypeName ?? string.Empty;
                var smplName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference idRef))
                {
                    var declarationTypeName = GetBaseTypeForDeclaration(idRef.Declaration);
                    newResult = _inspValueTreeFactory.Create(context.GetText(), declarationTypeName);
                }
            }

            if (newResult != null)
            {
                StoreVisitResult(context, newResult);
            }
        }

        private void Visit(VBAParser.LiteralExprContext context)
        {
            if (!ContextHasResult(context))
            {
                var nResult = _inspValueTreeFactory.Create(context.GetText());
                StoreVisitResult(context, nResult);
            }
        }

        private void VisitBinaryEvaluationContext(ParserRuleContext context)
        {
            VisitImpl(context);

            IUnreachableCaseInspectionValue nLHS = null;
            IUnreachableCaseInspectionValue nRHS = null;
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_contextValues.Keys.Contains(ctxt))
                {
                    if (nLHS is null)
                    {
                        nLHS = _contextValues[(ParserRuleContext)ctxt];
                    }
                    else if (nRHS is null)
                    {
                        nRHS = _contextValues[(ParserRuleContext)ctxt];
                    }
                }
            }

            //TODO: getting the opSymbol looks redundant 
            if (IsBinaryMathContext(context))
            {
                if (nLHS != null && nRHS != null && IsBinaryMathContext(context))
                {
                    var opSymbol = context.children.Where(ch => BinaryOps.Keys.Contains(ch.GetText())).First().GetText();
                    if (BinaryOps.ContainsKey(opSymbol))
                    {
                        var nResult = BinaryOps[opSymbol].Evaluate(nLHS, nRHS, EvaluationTypeName);
                        StoreVisitResult(context, nResult);
                    }
                }
            }
            else
            {
                if (nLHS != null && nRHS != null)
                {
                    var opSymbol = context.children.Where(ch => BinaryOps.Keys.Contains(ch.GetText())).First().GetText();
                    if (BinaryOps.ContainsKey(opSymbol))
                    {
                        var nResult = _inspValueTreeFactory.Create(context.GetText(), Tokens.Boolean);
                        StoreVisitResult(context, nResult);
                    }
                }
            }
        }

        private void Visit(VBAParser.UnaryMinusOpContext context)
        {
            VisitImpl(context);

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_contextValues.Keys.Contains(ctxt))
                {
                    var value = _contextValues[(ParserRuleContext)ctxt];
                    //var op = new UnreachableCaseInspectionMinusOp();
                    var op = UnaryOps[MathTokens.ADDITIVE_INVERSE];
                    var result = op.Evaluate(value, EvaluationTypeName);
                    StoreVisitResult(context, result);
                }
            }
        }

        private void StoreVisitResult(ParserRuleContext context, IUnreachableCaseInspectionValue inspValue)
        {
            if (ContextHasResult(context))
            {
                return;
            }
            var conformedInsp = new UnreachableCaseInspectionValueConformed(inspValue, inspValue.TypeName);
            _contextValues.Add(context, conformedInsp);
        }

        private bool ContextHasResult(ParserRuleContext context)
        {
            return _contextValues.ContainsKey(context);
        }

        private void VisitUnaryLogicalContext(ParserRuleContext context)
        {
            VisitImpl(context);

            IUnreachableCaseInspectionValue nLHS = null;
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count() && (nLHS is null); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_contextValues.Keys.Contains(ctxt))
                {
                    nLHS = _contextValues[(ParserRuleContext)ctxt];
                }
            }

           var opSymbol = context.children.Where(ch => UnaryOps.Keys.Contains(ch.GetText())).First().GetText();
            if (UnaryOps.ContainsKey(opSymbol))
            {
                var result = _inspValueTreeFactory.Create(context.GetText(), Tokens.Boolean);
                StoreVisitResult(context, result);
            }
        }

        private void VisitImpl(ParserRuleContext context)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext) && ch is ParserRuleContext).ToList();
            foreach (var ctxt in contextsOfInterest)
            {
                Visit((ParserRuleContext)ctxt);
            }
        }

        private void VisitSummaryContextType(ParserRuleContext parserRuleContext)
        {
            VisitImpl(parserRuleContext);
            var contextsOfInterest = parserRuleContext.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext) && ch is ParserRuleContext).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = (ParserRuleContext)contextsOfInterest[idx];
                if (_contextValues.Keys.Contains(ctxt))
                {
                    var value = _contextValues[ctxt];
                    StoreVisitResult(parserRuleContext, value);
                }
            }
        }

        private bool TryGetTheLExprValue(VBAParser.LExprContext ctxt, out string expressionValue, out string declaredTypeName)
        {
            expressionValue = string.Empty;
            declaredTypeName = string.Empty;
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
                            expressionValue = GetConstantDeclarationValueToken(dec);
                            declaredTypeName = dec.AsTypeIsBaseType ? dec.AsTypeName : dec.AsTypeDeclaration.AsTypeName;
                            return true;
                        }
                    }
                }
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
                        expressionValue = GetConstantDeclarationValueToken(declaration);
                        declaredTypeName = declaration.AsTypeName;
                        return true;
                    }
                }
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

        private string GetConstantDeclarationValueToken(Declaration valueDeclaration)
        {
            var contextsOfInterest = new List<ParserRuleContext>();
            var contexts = valueDeclaration.Context.children.ToList();
            var eqIndex = contexts.FindIndex(ch => ch.GetText().Equals(CompareTokens.EQ));
            for (int idx = eqIndex + 1; idx < contexts.Count(); idx++)
            {
                var childCtxt = contexts[idx];
                if (!(childCtxt is VBAParser.WhiteSpaceContext))
                {
                    contextsOfInterest.Add((ParserRuleContext)childCtxt);
                }
            }

            foreach (var child in contextsOfInterest)
            {
                Visit(child);
                if (_contextValues.TryGetValue(child, out IUnreachableCaseInspectionValue value))
                {
                    return value.ValueText;
                }
            }
            return string.Empty;
        }

        private static string GetBaseTypeForDeclaration(Declaration declaration)
        {
            if (!declaration.AsTypeIsBaseType)
            {
                return GetBaseTypeForDeclaration(declaration.AsTypeDeclaration);
            }
            return declaration.AsTypeName;
        }

        private static bool IsBinaryMathContext<T>(T child)
        {
            return child is VBAParser.MultOpContext
                || child is VBAParser.AddOpContext
                || child is VBAParser.PowOpContext
                || child is VBAParser.ModOpContext;
        }

        private static bool IsUnaryResultContext<T>(T child)
        {
            return child is VBAParser.SelectStartValueContext
                || child is VBAParser.SelectEndValueContext
                || child is VBAParser.ParenthesizedExprContext
                || child is VBAParser.SelectExpressionContext
                || child is VBAParser.RangeClauseContext;
        }

        private static bool IsLogicalContext<T>(T child)
        {
            return IsBinaryLogicalContext(child) || IsUnaryLogicalContext(child);
        }

        private static bool IsBinaryLogicalContext<T>(T child)
        {
            return child is VBAParser.RelationalOpContext
                || child is VBAParser.LogicalXorOpContext
                || child is VBAParser.LogicalAndOpContext
                || child is VBAParser.LogicalOrOpContext
                || child is VBAParser.LogicalEqvOpContext;
        }

        private static bool IsUnaryLogicalContext<T>(T child)
        {
            return child is VBAParser.LogicalNotOpContext;
        }
    }
}
