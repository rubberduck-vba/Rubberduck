using System;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPreprocessorVisitor : VBAConditionalCompilationParserBaseVisitor<IExpression>
    {
        private readonly SymbolTable<string, IValue> _symbolTable;
        private readonly ICharStream _stream;
        private readonly CommonTokenStream _tokenStream;

        public VBAPreprocessorVisitor(
            SymbolTable<string, IValue> symbolTable, 
            VBAPredefinedCompilationConstants predefinedConstants,
            Dictionary<string, short> userDefinedConstants,
            ICharStream stream,
            CommonTokenStream tokenStream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (tokenStream == null)
            {
                throw new ArgumentNullException(nameof(tokenStream));
            }
            if (symbolTable == null)
            {
                throw new ArgumentNullException(nameof(symbolTable));
            }
            if (predefinedConstants == null)
            {
                throw new ArgumentNullException(nameof(predefinedConstants));
            }

            _stream = stream;
            _tokenStream = tokenStream;
            _symbolTable = symbolTable;
            AddPredefinedConstantsToSymbolTable(predefinedConstants);
            AddUserDefinedConstantsToSymbolTable(userDefinedConstants);
        }

        private void AddPredefinedConstantsToSymbolTable(VBAPredefinedCompilationConstants predefinedConstants)
        {
            foreach (var constant in predefinedConstants.AllPredefinedConstants)
            {
                _symbolTable.AddOrUpdate(constant.Key, new DecimalValue(constant.Value));
            }
        }

        private void AddUserDefinedConstantsToSymbolTable(Dictionary<string, short> userDefinedConstants)
        {
            foreach (var constant in userDefinedConstants)
            {
                _symbolTable.AddOrUpdate(constant.Key, new DecimalValue(constant.Value));
            }
        }

        public override IExpression VisitCompilationUnit([NotNull] VBAConditionalCompilationParser.CompilationUnitContext context)
        {
            return Visit(context.ccBlock());
        }

        public override IExpression VisitPhysicalLine([NotNull] VBAConditionalCompilationParser.PhysicalLineContext context)
        {
            return new ConstantExpression(new TokensValue(context.GetTokens(_tokenStream)));
        }

        public override IExpression VisitCcBlock([NotNull] VBAConditionalCompilationParser.CcBlockContext context)
        {
            if (context.children == null)
            {
                return new ConstantExpression(EmptyValue.Value);
            }
            return new ConditionalCompilationBlockExpression(context.children.Select(child => Visit(child)).ToList());
        }

        public override IExpression VisitCcConst([NotNull] VBAConditionalCompilationParser.CcConstContext context)
        {
            return new ConditionalCompilationConstantExpression(
                    new ConstantExpression(new TokensValue(context.GetTokens(_tokenStream))),
                    new ConstantExpression(new StringValue(Identifier.GetName(context.ccVarLhs().name()))),
                    Visit(context.ccExpression()),
                    _symbolTable);
        }

        public override IExpression VisitCcIfBlock([NotNull] VBAConditionalCompilationParser.CcIfBlockContext context)
        {
            var ifCondTokens = new ConstantExpression(new TokensValue(context.ccIf().GetTokens( _tokenStream)));
            var ifCond = Visit(context.ccIf().ccExpression());
            var ifBlock = Visit(context.ccBlock());
            var elseIfCodeCondBlocks = context
                .ccElseIfBlock()
                .Select(elseIf =>
                        Tuple.Create<IExpression, IExpression, IExpression>(
                            new ConstantExpression(new TokensValue(elseIf.ccElseIf().GetTokens(_tokenStream))),
                            Visit(elseIf.ccElseIf().ccExpression()),
                            Visit(elseIf.ccBlock())))
                .ToList();

            IExpression elseCondTokens = null;
            IExpression elseBlock = null;
            if (context.ccElseBlock() != null)
            {
                elseCondTokens = new ConstantExpression(new TokensValue(context.ccElseBlock().ccElse().GetTokens(_tokenStream)));
                elseBlock = Visit(context.ccElseBlock().ccBlock());
            }
            var endIfTokens = new ConstantExpression(new TokensValue(context.ccEndIf().GetTokens(_tokenStream)));
            return new ConditionalCompilationIfExpression(
                    ifCondTokens,
                    ifCond,
                    ifBlock,
                    elseIfCodeCondBlocks,
                    elseCondTokens,
                    elseBlock,
                    endIfTokens);
        }

        private IExpression Visit(VBAConditionalCompilationParser.NameContext context)
        {
            return new NameExpression(
                new ConstantExpression(new StringValue(Identifier.GetName(context))),
                _symbolTable);
        }

        private IExpression VisitUnaryMinus(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            return new UnaryMinusExpression(Visit(context.ccExpression()[0]));
        }

        private IExpression VisitUnaryNot(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            return new UnaryNotExpression(Visit(context.ccExpression()[0]));
        }

        private IExpression VisitPlus(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            return new BinaryPlusExpression(Visit(context.ccExpression()[0]), Visit(context.ccExpression()[1]));
        }

        private IExpression VisitMinus(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            return new BinaryMinusExpression(Visit(context.ccExpression()[0]), Visit(context.ccExpression()[1]));
        }

        private IExpression Visit(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            if (context.literal() != null)
            {
                return Visit(context.literal());
            }
            else if (context.name() != null)
            {
                return Visit(context.name());
            }
            else if (context.LPAREN() != null)
            {
                return Visit(context.ccExpression()[0]);
            }
            else if (context.MINUS() != null && context.ccExpression().Length == 1)
            {
                return VisitUnaryMinus(context);
            }
            else if (context.NOT() != null)
            {
                return VisitUnaryNot(context);
            }
            else if (context.PLUS() != null)
            {
                return VisitPlus(context);
            }
            else if (context.MINUS() != null && context.ccExpression().Length == 2)
            {
                return VisitMinus(context);
            }
            else if (context.MULT() != null)
            {
                return VisitMult(context);
            }
            else if (context.DIV() != null)
            {
                return VisitDiv(context);
            }
            else if (context.INTDIV() != null)
            {
                return VisitIntDiv(context);
            }
            else if (context.MOD() != null)
            {
                return VisitMod(context);
            }
            else if (context.POW() != null)
            {
                return VisitPow(context);
            }
            else if (context.AMPERSAND() != null)
            {
                return VisitConcat(context);
            }
            else if (context.EQ() != null)
            {
                return VisitEq(context);
            }
            else if (context.NEQ() != null)
            {
                return VisitNeq(context);
            }
            else if (context.LT() != null)
            {
                return VisitLt(context);
            }
            else if (context.GT() != null)
            {
                return VisitGt(context);
            }
            else if (context.LEQ() != null)
            {
                return VisitLeq(context);
            }
            else if (context.GEQ() != null)
            {
                return VisitGeq(context);
            }
            else if (context.AND() != null)
            {
                return VisitAnd(context);
            }
            else if (context.OR() != null)
            {
                return VisitOr(context);
            }
            else if (context.XOR() != null)
            {
                return VisitXor(context);
            }
            else if (context.EQV() != null)
            {
                return VisitEqv(context);
            }
            else if (context.IMP() != null)
            {
                return VisitImp(context);
            }
            else if (context.IS() != null)
            {
                return VisitIs(context);
            }
            else if (context.LIKE() != null)
            {
                return VisitLike(context);
            }
            else
            {
                return VisitLibraryFunction(context);
            }
        }

        private IExpression VisitLibraryFunction(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var intrinsicFunction = context.intrinsicFunction();
            var functionName = intrinsicFunction.intrinsicFunctionName().GetText(_stream);
            var argument = Visit(intrinsicFunction.ccExpression());
            return VBALibrary.CreateLibraryFunction(functionName, argument);
        }

        private IExpression VisitLike(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var expr = Visit(context.ccExpression()[0]);
            var pattern = Visit(context.ccExpression()[1]);
            return new LikeExpression(expr, pattern);
        }

        private IExpression VisitIs(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new IsExpression(left, right);
        }

        private IExpression VisitImp(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalImpExpression(left, right);
        }

        private IExpression VisitEqv(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalEqvExpression(left, right);
        }

        private IExpression VisitXor(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalXorExpression(left, right);
        }

        private IExpression VisitOr(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalOrExpression(left, right);
        }

        private IExpression VisitAnd(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalAndExpression(left, right);
        }

        private IExpression VisitGeq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalGreaterOrEqualsExpression(left, right);
        }

        private IExpression VisitLeq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalLessOrEqualsExpression(left, right);
        }

        private IExpression VisitGt(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalGreaterThanExpression(left, right);
        }

        private IExpression VisitLt(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalLessThanExpression(left, right);
        }

        private IExpression VisitNeq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalNotEqualsExpression(left, right);
        }

        private IExpression VisitEq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new LogicalEqualsExpression(left, right);
        }

        private IExpression VisitConcat(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new ConcatExpression(left, right);
        }

        private IExpression VisitPow(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new PowExpression(left, right);
        }

        private IExpression VisitMod(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new ModExpression(left, right);
        }

        private IExpression VisitIntDiv(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new BinaryIntDivExpression(left, right);
        }

        private IExpression VisitMult(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new BinaryMultiplicationExpression(left, right);
        }

        private IExpression VisitDiv(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return new BinaryDivisionExpression(left, right);
        }

        private IExpression Visit(VBAConditionalCompilationParser.LiteralContext context)
        {
            if (context.HEXLITERAL() != null)
            {
                return VisitHexLiteral(context);
            }
            else if (context.OCTLITERAL() != null)
            {
                return VisitOctLiteral(context);
            }
            else if (context.DATELITERAL() != null)
            {
                return VisitDateLiteral(context);
            }
            else if (context.FLOATLITERAL() != null)
            {
                return VisitFloatLiteral(context);
            }
            else if (context.INTEGERLITERAL() != null)
            {
                return VisitIntegerLiteral(context);
            }
            else if (context.STRINGLITERAL() != null)
            {
                return VisitStringLiteral(context);
            }
            else if (context.TRUE() != null)
            {
                return new ConstantExpression(new BoolValue(true));
            }
            else if (context.FALSE() != null)
            {
                return new ConstantExpression(new BoolValue(false));
            }
            else if (context.NOTHING() != null || context.NULL() != null)
            {
                return new ConstantExpression(null);
            }
            else if (context.EMPTY() != null)
            {
                return new ConstantExpression(EmptyValue.Value);
            }
            throw new Exception(string.Format("Unexpected literal encountered: {0}", context.GetText(_stream)));
        }

        private IExpression VisitStringLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            return new StringLiteralExpression(new ConstantExpression(new StringValue(context.STRINGLITERAL().GetText())));
        }

        private IExpression VisitIntegerLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            return new NumberLiteralExpression(new ConstantExpression(new StringValue(context.INTEGERLITERAL().GetText())));
        }

        private IExpression VisitFloatLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            return new NumberLiteralExpression(new ConstantExpression(new StringValue(context.FLOATLITERAL().GetText())));
        }

        private IExpression VisitDateLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            return new DateLiteralExpression(new ConstantExpression(new StringValue(context.DATELITERAL().GetText())));
        }

        private IExpression VisitOctLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            return new OctNumberLiteralExpression(new ConstantExpression(new StringValue(context.OCTLITERAL().GetText())));
        }

        private IExpression VisitHexLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            return new HexNumberLiteralExpression(new ConstantExpression(new StringValue(context.HEXLITERAL().GetText())));
        }
    }
}
