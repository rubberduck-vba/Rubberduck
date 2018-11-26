using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing
{
    public static class ExpressionContextExtensions
    {
        public static bool IsLogicalContext(this VBAParser.ExpressionContext context)
        {
            return TryGetLogicalContextSymbol(context, out _);
        }

        public static bool IsMathContext(this VBAParser.ExpressionContext context)
        {
            return context.IsBinaryMathContext() || context.IsUnaryMathContext();
        }

        public static bool IsBinaryMathContext(this VBAParser.ExpressionContext context)
        {
            return context is VBAParser.MultOpContext   //MultOpContext includes both * and /
                || context is VBAParser.AddOpContext    //AddOpContet includes both + and -
                || context is VBAParser.PowOpContext
                || context is VBAParser.IntDivOpContext
                || context is VBAParser.ModOpContext;
        }

        public static bool IsUnaryMathContext(this VBAParser.ExpressionContext context)
        {
            return context is VBAParser.UnaryMinusOpContext;
        }

        public static bool IsBinaryLogicalContext(this VBAParser.ExpressionContext context)
        {
            return context is VBAParser.RelationalOpContext
                || context is VBAParser.LogicalXorOpContext
                || context is VBAParser.LogicalAndOpContext
                || context is VBAParser.LogicalOrOpContext
                || context is VBAParser.LogicalImpOpContext
                || context is VBAParser.LogicalEqvOpContext;
        }

        public static bool IsUnaryLogicalContext(this VBAParser.ExpressionContext context)
        {
            return context is VBAParser.LogicalNotOpContext;
        }

        public static bool TryGetLogicalContextSymbol(this VBAParser.ExpressionContext context, out string symbol)
        {
            if (context.TryGetRelationalOpContextSymbol(out symbol))
            {
                return true;
            }

            switch (context)
            {
                case VBAParser.LogicalXorOpContext _:
                    symbol = context.GetToken(VBAParser.XOR, 0).GetText();
                    return true;
                case VBAParser.LogicalAndOpContext _:
                    symbol = context.GetToken(VBAParser.AND, 0).GetText();
                    return true;
                case VBAParser.LogicalOrOpContext _:
                    symbol = context.GetToken(VBAParser.OR, 0).GetText();
                    return true;
                case VBAParser.LogicalEqvOpContext _:
                    symbol = context.GetToken(VBAParser.EQV, 0).GetText();
                    return true;
                case VBAParser.LogicalImpOpContext _:
                    symbol = context.GetToken(VBAParser.IMP, 0).GetText();
                    return true;
                case VBAParser.LogicalNotOpContext _:
                    symbol = context.GetToken(VBAParser.NOT, 0).GetText();
                    return true;
                default:
                    symbol = string.Empty;
                    return false;
            }
        }

        public static bool TryGetRelationalOpContextSymbol(this VBAParser.ExpressionContext context, out string opSymbol)
        {
            switch (context)
            {
                case VBAParser.RelationalOpContext ctxt:
                    var terminalNode = ctxt.EQ() ?? ctxt.GEQ() ?? ctxt.GT() ?? ctxt.LEQ()
                        ?? ctxt.LIKE() ?? ctxt.LT() ?? ctxt.NEQ();
                    opSymbol = terminalNode.GetText();
                    return true;
                default:
                    opSymbol = string.Empty;
                    return false;
            }
        }
    }
}
