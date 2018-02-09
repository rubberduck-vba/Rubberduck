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
    public class ContextValueVisitor : IParseTreeVisitor<ParseTreeValue>
    {
        private Dictionary<ParserRuleContext, ParseTreeValue> _valueResolvedContexts;
        private Dictionary<ParserRuleContext, ParseTreeValue> _unResolvedContexts;
        private RubberduckParserState _state;

        private static bool IsBinaryMathContext<T>(T child)
        {
            return child is VBAParser.MultOpContext
                || child is VBAParser.AddOpContext
                || child is VBAParser.PowOpContext
                || child is VBAParser.ModOpContext;
        }

        private static bool IsUnaryOpContext<T>(T child)
        {
            return
                //child is VBAParser.SelectCaseStmtContext
                 //child is VBAParser.SelectExpressionContext
                //|| child is VBAParser.CaseClauseContext
                //|| child is VBAParser.RangeClauseContext
                 child is VBAParser.SelectStartValueContext
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
        }

        public ContextValueVisitor(RubberduckParserState state, string evaluationTypeName)
        {
            _state = state;
            _valueResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            _unResolvedContexts = new Dictionary<ParserRuleContext, ParseTreeValue>();
            EvaluationTypeName = evaluationTypeName ?? string.Empty;
        }

        public bool IsCompareToken(string token) => AlgebraicLogicInversions.Keys.Contains(token);
        public string EvaluationTypeName { set; get; } = string.Empty;
        public RubberduckParserState State => _state;
        public Dictionary<ParserRuleContext, ParseTreeValue> ValueResolvedContexts => _valueResolvedContexts;
        public Dictionary<ParserRuleContext, ParseTreeValue> UnresolvedContexts => _unResolvedContexts;
        public Dictionary<ParserRuleContext, long> ResolvedContextsAsLongs => ConvertToLong(_valueResolvedContexts);
        public Dictionary<ParserRuleContext, bool> ResolvedContextsAsBooleans => ConvertToBoolean(_valueResolvedContexts);
        public Dictionary<ParserRuleContext, double> ResolvedContextsAsDoubles => ConvertToDouble(_valueResolvedContexts);
        public Dictionary<ParserRuleContext, decimal> ResolvedContextsAsCurrency => ConvertToCurrency(_valueResolvedContexts);
        public Dictionary<ParserRuleContext, byte> ResolvedContextsAsBytes => ConvertToByte(_valueResolvedContexts);
        public Dictionary<ParserRuleContext, string> ResolvedContextsAsStrings => ConvertToString(_valueResolvedContexts);


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

        public ContextValueResults<long> ResultsAsLong()
        {
            var converted = ConvertToLong(_valueResolvedContexts);
            return new ContextValueResults<long>(converted, _unResolvedContexts)
            {
                EvaluationTypeName = EvaluationTypeName
            };
        }

        public ContextValueResults<bool> ResultsAsBoolean()
        {
            var converted = ConvertToBoolean(_valueResolvedContexts);
            return new ContextValueResults<bool>(converted, _unResolvedContexts)
            {
                EvaluationTypeName = EvaluationTypeName
            };
        }

        public ContextValueResults<byte> ResultsAsByte()
        {
            var converted = ConvertToByte(_valueResolvedContexts);
            return new ContextValueResults<byte>(converted, _unResolvedContexts)
            {
                EvaluationTypeName = EvaluationTypeName
            };
        }

        public ContextValueResults<string> ResultsAsString()
        {
            var converted = ConvertToString(_valueResolvedContexts);
            return new ContextValueResults<string>(converted, _unResolvedContexts)
            {
                EvaluationTypeName = EvaluationTypeName
            };
        }

        public ContextValueResults<double> ResultsAsDouble()
        {
            var converted = ConvertToDouble(_valueResolvedContexts);
            return new ContextValueResults<double>(converted, _unResolvedContexts)
            {
                EvaluationTypeName = EvaluationTypeName
            };
        }

        public ContextValueResults<decimal> ResultsAsCurrency()
        {
            var converted = ConvertToCurrency(_valueResolvedContexts);
            return new ContextValueResults<decimal>(converted, _unResolvedContexts)
            {
                EvaluationTypeName = EvaluationTypeName
            };
        }

        private static Dictionary<ParserRuleContext, long> ConvertToLong(Dictionary<ParserRuleContext, ParseTreeValue> valueResolvedContexts)
        {
            var converted = new Dictionary<ParserRuleContext, long>();
            foreach (var key in valueResolvedContexts.Keys)
            {
                if (valueResolvedContexts[key].TryGetValue(out long val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        private static Dictionary<ParserRuleContext, bool> ConvertToBoolean(Dictionary<ParserRuleContext, ParseTreeValue> valueResolvedContexts)
        {
            var converted = new Dictionary<ParserRuleContext, bool>();
            foreach (var key in valueResolvedContexts.Keys)
            {
                if (valueResolvedContexts[key].TryGetValue(out bool val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        private static Dictionary<ParserRuleContext, byte> ConvertToByte(Dictionary<ParserRuleContext, ParseTreeValue> valueResolvedContexts)
        {
            var converted = new Dictionary<ParserRuleContext, byte>();
            foreach (var key in valueResolvedContexts.Keys)
            {
                if (valueResolvedContexts[key].TryGetValue(out byte val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        private static Dictionary<ParserRuleContext, string> ConvertToString(Dictionary<ParserRuleContext, ParseTreeValue> valueResolvedContexts)
        {
            var converted = new Dictionary<ParserRuleContext, string>();
            foreach (var key in valueResolvedContexts.Keys)
            {
                if (valueResolvedContexts[key].TryGetValue(out string val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        private static Dictionary<ParserRuleContext, double> ConvertToDouble(Dictionary<ParserRuleContext, ParseTreeValue> valueResolvedContexts)
        {
            var converted = new Dictionary<ParserRuleContext, double>();
            foreach (var key in valueResolvedContexts.Keys)
            {
                if (valueResolvedContexts[key].TryGetValue(out double val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        private static Dictionary<ParserRuleContext, decimal> ConvertToCurrency(Dictionary<ParserRuleContext, ParseTreeValue> valueResolvedContexts)
        {
            var converted = new Dictionary<ParserRuleContext, decimal>();
            foreach (var key in valueResolvedContexts.Keys)
            {
                if (valueResolvedContexts[key].TryGetValue(out decimal val))
                {
                    converted.Add(key, val);
                }
            }
            return converted;
        }

        //Used to modify logic operators to convert LHS and RHS for expressions like '5 > x' (=> 'x < 5')
        private static Dictionary<string, string> AlgebraicLogicInversions = new Dictionary<string, string>()
        {
            [CompareTokens.EQ] = CompareTokens.EQ,
            [CompareTokens.NEQ] = CompareTokens.NEQ,
            [CompareTokens.LT] = CompareTokens.GT,
            [CompareTokens.LTE] = CompareTokens.GTE,
            [CompareTokens.GT] = CompareTokens.LT,
            [CompareTokens.GTE] = CompareTokens.LTE
        };

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

        //public void Accept(ISelectStmtParserTreeVisitor visitor)
        //{
        //    throw new NotImplementedException();
        //}

        public virtual ParseTreeValue Visit(IParseTree tree)
        {
            if(tree is ParserRuleContext prCtxt)
            {
                Visit(prCtxt);
                if(this._valueResolvedContexts.TryGetValue(prCtxt, out ParseTreeValue value))
                {
                    return value;
                }
            }
            return ParseTreeValue.Null;
            //throw new NotImplementedException();
        }

        public virtual ParseTreeValue VisitChildren(IRuleNode node)
        {
            if (node is ParserRuleContext prCtxt)
            {
                Visit(prCtxt);
                if (this._unResolvedContexts.TryGetValue(prCtxt, out ParseTreeValue value))
                {
                    return value;
                }
            }
            return ParseTreeValue.Null;
            //throw new NotImplementedException();
        }

        public virtual ParseTreeValue VisitTerminal(ITerminalNode node)
        {
            throw new NotImplementedException();
        }

        public virtual ParseTreeValue VisitErrorNode(IErrorNode node)
        {
            throw new NotImplementedException();
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
            if (IsUnaryOpContext(parserRuleContext))
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
                Visit(sCSC);
            }
            else if (parserRuleContext is VBAParser.SelectExpressionContext sEC)
            {
                Visit(sEC);
            }
            else if (parserRuleContext is VBAParser.CaseClauseContext cCC)
            {
                Visit(cCC);
            }
            else if (parserRuleContext is VBAParser.RangeClauseContext rCC)
            {
                Visit(rCC);
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

        public void Visit(VBAParser.SelectCaseStmtContext context)
        {
            VisitImpl(context);
            //if (!ContextHasResult(context))
            //{
            //    var selectExpr = context.selectExpression();
            //    var result = new ParseTreeValue(selectExpr.GetText());
            //    if (!EvaluationTypeName.Equals(string.Empty))
            //    {
            //        result.UseageTypeName = EvaluationTypeName;
            //    }
            //    else
            //    {
            //        result.UseageTypeName = result.DerivedTypeName;
            //    }

            //    //StoreVisitResult(context, result);
            //}
        }

        public void Visit(VBAParser.SelectExpressionContext context)
        {
            VisitImpl(context);
            //if (!ContextHasResult(context))
            //{
            //    var result = new ParseTreeValue(context.GetText());
            //    if (!EvaluationTypeName.Equals(string.Empty))
            //    {
            //        result.UseageTypeName = EvaluationTypeName;
            //    }
            //    else
            //    {
            //        result.UseageTypeName = result.DerivedTypeName;
            //    }

            //    //StoreVisitResult(context, result);
            //}
        }

        public void Visit(VBAParser.CaseClauseContext context)
        {
            VisitImpl(context);
            //if (!ContextHasResult(context))
            //{
            //    var text = string.Empty;
            //    foreach ( var rg in context.rangeClause())
            //    {
            //        if(text.Length > 0)
            //        {
            //            text = $"{text},{rg.GetText()}";
            //        }
            //        else
            //        {
            //            text = $"{rg.GetText()}";
            //        }
            //    }
            //    var result = new ParseTreeValue(text);
            //    if (!EvaluationTypeName.Equals(string.Empty))
            //    {
            //        result.UseageTypeName = EvaluationTypeName;
            //    }
            //    else
            //    {
            //        result.UseageTypeName = result.DerivedTypeName;
            //    }

            //    //StoreVisitResult(context, result);
            //}
        }

        public void Visit(VBAParser.RangeClauseContext context)
        {
            VisitImpl(context);
            //if (!ContextHasResult(context))
            //{
            //    var result = new ParseTreeValue(context.GetText());
            //    if (!EvaluationTypeName.Equals(string.Empty))
            //    {
            //        result.UseageTypeName = EvaluationTypeName;
            //    }
            //    else
            //    {
            //        result.UseageTypeName = result.DerivedTypeName;
            //    }

            //    //StoreVisitResult(context, result);
            //}
        }

        public void Visit(VBAParser.LExprContext context)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            ParseTreeValue result = null;
            //var lexprTypeName = this.EvaluationTypeName ?? string.Empty;
            if (TryGetTheLExprValue(context, out string lexprValue, out string lexprTypeName))
            {
                //result = lexprTypeName.Length > 0 ? new ParseTreeValue(lexprValue, lexprTypeName) : new ParseTreeValue(lexprValue, lexprTypeName);
                result = new ParseTreeValue(lexprValue, lexprTypeName); // : new ParseTreeValue(lexprValue, lexprTypeName);
            }
            else
            {
                var smplNameExprTypeName = this.EvaluationTypeName ?? string.Empty;
                var smplName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference idRef))
                {
                    var theTypeName = GetBaseTypeForDeclaration(idRef.Declaration);
                    result = new ParseTreeValue(context.GetText(), theTypeName);
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

            var contextsOfInterest = context.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)).ToList();
            ParseTreeValue LHS = null;
            ParseTreeValue RHS = null;
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

        //private bool TryGetTheLExprValueX(VBAParser.LExprContext ctxt, out string expressionValue, ref string typeName)
        //{
        //    expressionValue = string.Empty;
        //    if (ctxt.TryGetChildContext(out VBAParser.MemberAccessExprContext memberAccess))
        //    {
        //        var member = memberAccess.GetChild<VBAParser.UnrestrictedIdentifierContext>();

        //        if (TryGetIdentifierReferenceForContext(member, out IdentifierReference idRef))
        //        {
        //            var dec = idRef.Declaration;
        //            if (dec.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
        //            {
        //                var theCtxt = dec.Context;
        //                if (theCtxt is VBAParser.EnumerationStmt_ConstantContext)
        //                {
        //                    expressionValue = GetConstantDeclarationValue(dec);
        //                    typeName = dec.AsTypeIsBaseType ? dec.AsTypeName : dec.AsTypeDeclaration.AsTypeName;
        //                    return true;
        //                }
        //            }
        //        }
        //        return false;
        //    }

        //    if (ctxt.TryGetChildContext(out VBAParser.SimpleNameExprContext smplName))
        //    {
        //        if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference rangeClauseIdentifierReference))
        //        {
        //            var declaration = rangeClauseIdentifierReference.Declaration;
        //            if (declaration.DeclarationType.HasFlag(DeclarationType.Constant)
        //                || declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
        //            {
        //                expressionValue = GetConstantDeclarationValue(declaration);
        //                typeName = declaration.AsTypeName;
        //                return true;
        //            }
        //        }
        //    }
        //    return false;
        //}

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
                            expressionValue = GetConstantDeclarationValue(dec);
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
                        expressionValue = GetConstantDeclarationValue(declaration);
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

        private string GetConstantDeclarationValue(Declaration valueDeclaration)
        {
            var contextsOfInterest = GetRHSContexts(valueDeclaration.Context.children.ToList());
            foreach (var child in contextsOfInterest)
            {
                Visit(child);
                if (ContextHasResult(child))
                {
                    return ValueResolvedContexts[child].AsString();
                }
            }
            return string.Empty;
        }

        private List<ParserRuleContext> GetRHSContexts(List<IParseTree> contexts)
        {
            var contextsOfInterest = new List<ParserRuleContext>();
            //TODO: what happens if COmpareTokens.EQ is not found
            //var eqIndex = contexts.FindIndex(ch => ch.GetText().Equals(CompareTokens.EQ));
            var eqIndex = contexts.FindIndex(ch => IsCompareToken(ch.GetText()));
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
    }
}
