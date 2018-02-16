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
    public interface IParseTreeValueResults
    {
        IParseTreeValueResults Add(IParseTreeValueResults ptValues);
        Dictionary<ParserRuleContext, ParseTreeValue> ValueResolvedContexts { set; get; }
        Dictionary<ParserRuleContext, ParseTreeValue> VariableContexts { set; get; }
        List<ParserRuleContext> AllContexts { get; }
        //List<ParseTreeValue> AllValues { get; }
        //List<ParseTreeValue> Constants { get; }
        //List<ParseTreeValue> Variables { get; }
        List<ParseTreeValue> RangeClauseResults();
        ParseTreeValue Result(ParserRuleContext context);
        Dictionary<ParserRuleContext, long> ValueResultsAsLong();
        Dictionary<ParserRuleContext, double> ValueResultsAsDouble();
        Dictionary<ParserRuleContext, decimal> ValueResultsAsDecimal();
        Dictionary<ParserRuleContext, bool> ValueResultsAsBoolean();
        Dictionary<ParserRuleContext, string> ValueResultsAsString();
    }

    public class ParseTreeValueResults : IParseTreeValueResults
    {
        private Dictionary<ParserRuleContext, ParseTreeValue> _valueResolvedContexts;
        private Dictionary<ParserRuleContext, ParseTreeValue> _unResolvedContexts;

        public ParseTreeValueResults()
        {
            _valueResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            _unResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
        }

        public ParseTreeValueResults(Dictionary<ParserRuleContext, ParseTreeValue> valueResolved, Dictionary<ParserRuleContext, ParseTreeValue> unResolved)
        {
            _valueResolvedContexts = valueResolved;
            _unResolvedContexts = unResolved;
        }

        public IParseTreeValueResults Add(IParseTreeValueResults ptValues)
        {
            foreach (var ptVal in ptValues.ValueResolvedContexts)
            {
                _valueResolvedContexts.Add(ptVal.Key, ptVal.Value);
            }
            foreach (var ptVal in ptValues.VariableContexts)
            {
                _unResolvedContexts.Add(ptVal.Key, ptVal.Value);
            }
            return this;
        }


        //public List<ParseTreeValue> AllValues
        //{
        //    get
        //    {
        //        var all = new List<ParseTreeValue>();
        //        all.AddRange(Constants);
        //        all.AddRange(Variables);
        //        return all;
        //    }
        //}

        public List<ParserRuleContext> AllContexts
        {
            get
            {
                var all = new List<ParserRuleContext>();
                all.AddRange(ValueResolvedContexts.Keys);
                all.AddRange(VariableContexts.Keys);
                return all;
            }
        }
        //public List<ParseTreeValue> Constants => ValueResolvedContexts.Values.ToList();
        //public List<ParseTreeValue> Variables => VariableContexts.Values.ToList();
        public Dictionary<ParserRuleContext, ParseTreeValue> ValueResolvedContexts { get => _valueResolvedContexts; set => _valueResolvedContexts = value; }
        public Dictionary<ParserRuleContext, ParseTreeValue> VariableContexts { get => _unResolvedContexts; set => _unResolvedContexts = value; }
        public IEnumerable<ParseTreeValue> VariablesForContexts(List<ParserRuleContext> contexts)
        {
            return VariableContexts.Where(vrc => contexts.Contains(vrc.Key)).Select(r => r.Value);
        }

        public  List<ParseTreeValue> RangeClauseResults()
        {
            var rangeClauseContexts = AllContexts.Where(ac => ac.IsDescendentOf<VBAParser.CaseClauseContext>());
            var results = new List<ParseTreeValue>();
            foreach(var context in rangeClauseContexts)
            {
                results.Add(Result(context));
            }
            return results;
        }

        public ParseTreeValue Result(ParserRuleContext context)
        {
            if (ValueResolvedContexts.ContainsKey(context))
            {
                return ValueResolvedContexts[context];
            }
            else if (VariableContexts.ContainsKey(context))
            {
                return VariableContexts[context];
            }
            return ParseTreeValue.Null;
        }

        public Dictionary<ParserRuleContext, long> ValueResultsAsLong()
        {
            var converted = new Dictionary<ParserRuleContext, long>();
            foreach (var key in _valueResolvedContexts.Keys)
            {
                if (_valueResolvedContexts[key].TryGetValue(out long val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        public Dictionary<ParserRuleContext, double> ValueResultsAsDouble()
        {
            var converted = new Dictionary<ParserRuleContext, double>();
            foreach (var key in _valueResolvedContexts.Keys)
            {
                if (_valueResolvedContexts[key].TryGetValue(out double val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        public Dictionary<ParserRuleContext, decimal> ValueResultsAsDecimal()
        {
            var converted = new Dictionary<ParserRuleContext, decimal>();
            foreach (var key in _valueResolvedContexts.Keys)
            {
                if (_valueResolvedContexts[key].TryGetValue(out decimal val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        public Dictionary<ParserRuleContext, bool> ValueResultsAsBoolean()
        {
            var converted = new Dictionary<ParserRuleContext, bool>();
            foreach (var key in _valueResolvedContexts.Keys)
            {
                if (_valueResolvedContexts[key].TryGetValue(out bool val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        public Dictionary<ParserRuleContext, string> ValueResultsAsString()
        {
            var converted = new Dictionary<ParserRuleContext, string>();
            foreach (var key in _valueResolvedContexts.Keys)
            {
                if (_valueResolvedContexts.ContainsKey(key))
                {
                    converted.Add(key, _valueResolvedContexts[key].ToString());
                }
            }
            return converted;
        }
    }

    public class ContextValueVisitor : IParseTreeVisitor<IParseTreeValueResults>
    {
        private Dictionary<ParserRuleContext, ParseTreeValue> _valueResolvedContexts;
        private Dictionary<ParserRuleContext, ParseTreeValue> _unResolvedContexts;
        private RubberduckParserState _state;
        private IRuleNode _lastVisitStartNode;

        private static bool IsBinaryMathContext<T>(T child)
        {
            return child is VBAParser.MultOpContext
                || child is VBAParser.AddOpContext
                || child is VBAParser.PowOpContext
                || child is VBAParser.ModOpContext;
        }

        private static bool IsUnaryValueContext<T>(T child)
        {
            return child is VBAParser.SelectStartValueContext
                || child is VBAParser.SelectEndValueContext
                || child is VBAParser.ParenthesizedExprContext;
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

        public ContextValueVisitor(RubberduckParserState state)
        {
            _state = state;
            _valueResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            _unResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            _lastVisitStartNode = null;
        }

        public ContextValueVisitor(RubberduckParserState state, string evaluationTypeName)
        {
            _state = state;
            _valueResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            _unResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            EvaluationTypeName = evaluationTypeName ?? string.Empty;
            _lastVisitStartNode = null;
        }

        public string EvaluationTypeName { set; get; } = string.Empty;
        private RubberduckParserState State => _state;
        public Dictionary<ParserRuleContext, ParseTreeValue> ValueResolvedContexts => _valueResolvedContexts;
        public Dictionary<ParserRuleContext, ParseTreeValue> UnresolvedContexts => _unResolvedContexts;


        public ParseTreeValue ContextValue(ParserRuleContext prCtxt)
        {
            if(ValueResolvedContexts.TryGetValue(prCtxt, out ParseTreeValue resolvedValue))
            {
                return resolvedValue;
            }
            else if(UnresolvedContexts.TryGetValue(prCtxt, out ParseTreeValue unresolvedValue))
            {
                return unresolvedValue;
            }
            return ParseTreeValue.Null;
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

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>
            BinaryMathOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
            {
                [MathTokens.ADD] = (LHS, RHS) => LHS + RHS,
                [MathTokens.SUBTRACT] = (LHS, RHS) => LHS - RHS,
                [MathTokens.MULT] = (LHS, RHS) => LHS * RHS,
                [MathTokens.DIV] = (LHS, RHS) => LHS / RHS,
                [MathTokens.POW] = (LHS, RHS) => ParseTreeValue.Pow(LHS, RHS),
                [MathTokens.MOD] = (LHS, RHS) => LHS % RHS
            };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>
            BinaryLogicalOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue, ParseTreeValue>>()
            {
                [CompareTokens.GT] = (LHS, RHS) => LHS > RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.GTE] = (LHS, RHS) => LHS >= RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.LT] = (LHS, RHS) => LHS < RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.LTE] = (LHS, RHS) => LHS <= RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.EQ] = (LHS, RHS) => LHS == RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [CompareTokens.NEQ] = (LHS, RHS) => LHS != RHS ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.And] = (LHS, RHS) => LHS.AsBoolean().Value && RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.Or] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.XOr] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False,
                [Tokens.Not] = (LHS, RHS) => LHS.AsBoolean().Value || RHS.AsBoolean().Value ? ParseTreeValue.True : ParseTreeValue.False
                //["Eqv"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False
                //["Imp"] = (LHS, RHS) => LHS.AsBoolean().Value ^ RHS.AsBoolean().Value ? UnreachableCaseInspectionValue.True : UnreachableCaseInspectionValue.False,
            };

        private static Dictionary<string, Func<ParseTreeValue, ParseTreeValue>>
            UnaryLogicalOps = new Dictionary<string, Func<ParseTreeValue, ParseTreeValue>>()
            {
                [Tokens.Not] = (LHS) => !(LHS.AsBoolean().Value) ? ParseTreeValue.True : ParseTreeValue.False
            };

        private void StoreVisitResult(ParserRuleContext context, ParseTreeValue result)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            if (result.HasValue)
            {
                _valueResolvedContexts.Add(context, result);
            }
            else
            {
                _unResolvedContexts.Add(context, result);
            }
        }

        private bool ContextHasResult(ParserRuleContext context)
        {
            return _valueResolvedContexts.Keys.Contains(context) || _unResolvedContexts.Keys.Contains(context);
        }

        public IParseTreeValueResults ReVisitUsingType(string evalTypeName, IRuleNode node = null)
        {
            EvaluationTypeName = evalTypeName;
            _unResolvedContexts.Clear();
            _valueResolvedContexts.Clear();

            if(node is null)
            {
                return VisitChildren(_lastVisitStartNode);
            }
            else
            {
                _lastVisitStartNode = node;
                return VisitChildren(node);
            }
        }

        public virtual IParseTreeValueResults Visit(IParseTree tree)
        {
            if (tree is ParserRuleContext prCtxt)
            {
                if(!(prCtxt is VBAParser.WhiteSpaceContext))
                {
                    Visit(prCtxt);
                    return new ParseTreeValueResults(_valueResolvedContexts, _unResolvedContexts);
                }
            }
            return new ParseTreeValueResults();
        }

        public virtual IParseTreeValueResults VisitChildren(IRuleNode node)
        {
            _lastVisitStartNode = node;
            for (var idx = 0; idx < node.ChildCount; idx++)
            {
                var child = node.GetChild(idx);
                Visit(child);
            }
            return new ParseTreeValueResults(_valueResolvedContexts, _unResolvedContexts);
        }

        public virtual IParseTreeValueResults VisitTerminal(ITerminalNode node)
        {
            return new ParseTreeValueResults();
        }

        public virtual IParseTreeValueResults VisitErrorNode(IErrorNode node)
        {
            return new ParseTreeValueResults();
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

        public void Visit(ParserRuleContext parserRuleContext)
        {
            if (IsUnaryValueContext(parserRuleContext))
            {
                VisitUnaryContext(parserRuleContext);
            }
            else if (parserRuleContext is VBAParser.LExprContext lExpr)
            {
                Visit(lExpr);
            }
            else if (parserRuleContext is VBAParser.LiteralExprContext litExpr)
            {
                Visit(litExpr);
            }
            else if (parserRuleContext is VBAParser.SelectCaseStmtContext sCSC)
            {
                VisitImpl(sCSC);
            }
            else if (parserRuleContext is VBAParser.SelectExpressionContext sEC)
            {
                VisitImpl(sEC);
            }
            else if (parserRuleContext is VBAParser.CaseClauseContext cCC)
            {
                VisitImpl(cCC);
            }
            else if (parserRuleContext is VBAParser.RangeClauseContext rCC)
            {
                VisitImpl(rCC);
            }
            else if (IsBinaryMathContext(parserRuleContext))
            {
                VisitBinaryMathContext(parserRuleContext);
            }
            else if (IsLogicalContext(parserRuleContext))
            {
                VisitLogicalContext(parserRuleContext);
            }
            else if (parserRuleContext is VBAParser.UnaryMinusOpContext uMinusOp)
            {
                Visit(uMinusOp);
            }
        }

        public void Visit(VBAParser.LExprContext context)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            ParseTreeValue result = null;
            if (TryGetTheLExprValue(context, out string lexprValue, out string declaredTypeName))
            {
                result = new ParseTreeValue(lexprValue, declaredTypeName);
            }
            else
            {
                var smplNameExprTypeName = this.EvaluationTypeName ?? string.Empty;
                var smplName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference idRef))
                {
                    var declarationTypeName = GetBaseTypeForDeclaration(idRef.Declaration);
                    result = new ParseTreeValue(context.GetText(), declarationTypeName);
                    if (!smplNameExprTypeName.Equals(string.Empty))
                    {
                        result.UseageTypeName = smplNameExprTypeName;
                    }
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
                var result = new ParseTreeValue(context.GetText());
                if(!EvaluationTypeName.Equals(string.Empty))
                {
                    result.UseageTypeName = EvaluationTypeName;
                }
                else
                {
                    result.UseageTypeName = result.DerivedTypeName;
                }

                StoreVisitResult(context, result);
            }
        }

        public void Visit(VBAParser.UnaryMinusOpContext context)
        {
            VisitImpl(context);
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_valueResolvedContexts.Keys.Contains(ctxt) && _valueResolvedContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    StoreVisitResult(context, _valueResolvedContexts[(ParserRuleContext)ctxt].AdditiveInverse);
                }
            }
        }

        public void VisitBinaryMathContext(ParserRuleContext context)
        {
            VisitImpl(context);

            ParseTreeValue LHS = null;
            ParseTreeValue RHS = null;
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_valueResolvedContexts.Keys.Contains(ctxt) && _valueResolvedContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    if (LHS is null)
                    {
                        LHS = _valueResolvedContexts[(ParserRuleContext)ctxt];
                    }
                    else if (RHS is null)
                    {
                        RHS = _valueResolvedContexts[(ParserRuleContext)ctxt];
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

        public void VisitLogicalContext(ParserRuleContext context)
        {
            VisitImpl(context);

            var isBinaryCtxt = IsBinaryLogicalContext(context);

            ParseTreeValue LHS = null;
            ParseTreeValue RHS = null;
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_valueResolvedContexts.Keys.Contains(ctxt) && _valueResolvedContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    if (LHS is null)
                    {
                        LHS = _valueResolvedContexts[(ParserRuleContext)ctxt];
                    }
                    else if (RHS is null && isBinaryCtxt)
                    {
                        RHS = _valueResolvedContexts[(ParserRuleContext)ctxt];
                    }
                }
            }

            if (isBinaryCtxt)
            {
                if (LHS != null && LHS.HasValue && RHS != null && RHS.HasValue)
                {
                    var opSymbol = context.children.Where(ch => BinaryLogicalOps.Keys.Contains(ch.GetText())).First().GetText();
                    if (BinaryLogicalOps.ContainsKey(opSymbol))
                    {
                        var result = new ParseTreeValue(BinaryLogicalOps[opSymbol](LHS, RHS).ToString(), EvaluationTypeName);
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
                        var result = new ParseTreeValue(UnaryLogicalOps[opSymbol](LHS).ToString(), EvaluationTypeName);
                        StoreVisitResult(context, result);
                    }
                }
            }
        }

        private void VisitUnaryContext(ParserRuleContext context)
        {
            VisitImpl(context);
            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            var num = contextsOfInterest.Count;
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest[idx];
                if (_valueResolvedContexts.Keys.Contains(ctxt) && _valueResolvedContexts[(ParserRuleContext)ctxt].HasValue)
                {
                    StoreVisitResult(context, _valueResolvedContexts[(ParserRuleContext)ctxt]);
                }
                else if(_unResolvedContexts.Keys.Contains(ctxt))// && _valueResolvedContexts[(ParserRuleContext)ctxt])
                {
                    StoreVisitResult(context, _unResolvedContexts[(ParserRuleContext)ctxt]);
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
                if(ValueResolvedContexts.TryGetValue(child, out ParseTreeValue value))
                {
                    return value.ToString();
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
    }
}
