using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIValueExpressionEvaluator
    {
        IUCIValue Evaluate(IUCIValue LHS, IUCIValue RHS, string opSymbol);
        IUCIValue Evaluate(IUCIValue LHS, string opSymbol);
    }

    public class UCIValueExpressionEvaluator : IUCIValueExpressionEvaluator
    {
        private readonly IUCIValueFactory _valueFactory;

        private static Dictionary<string, Func<double, double, double>> MathOpsBinary = new Dictionary<string, Func<double, double, double>>()
        {
            [MathTokens.MULT] = delegate (double LHS, double RHS) { return LHS * RHS; },
            [MathTokens.DIV] = delegate (double LHS, double RHS) { return LHS / RHS; },
            [MathTokens.ADD] = delegate (double LHS, double RHS) { return LHS + RHS; },
            [MathTokens.SUBTRACT] = delegate (double LHS, double RHS) { return LHS - RHS; },
            [MathTokens.POW] = Math.Pow,
            [MathTokens.MOD] = delegate (double LHS, double RHS) { return LHS % RHS; },
        };

        private static Dictionary<string, Func<double, double, bool>> LogicOpsBinary = new Dictionary<string, Func<double, double, bool>>()
        {
            [CompareTokens.EQ] = delegate (double LHS, double RHS) { return LHS == RHS; },
            [CompareTokens.NEQ] = delegate (double LHS, double RHS) { return LHS != RHS; },
            [CompareTokens.LT] = delegate (double LHS, double RHS) { return LHS < RHS; },
            [CompareTokens.LTE] = delegate (double LHS, double RHS) { return LHS <= RHS; },
            [CompareTokens.GT] = delegate (double LHS, double RHS) { return LHS > RHS; },
            [CompareTokens.GTE] = delegate (double LHS, double RHS) { return LHS >= RHS; },
            [Tokens.And] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) && Convert.ToBoolean(RHS); },
            [Tokens.Or] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) || Convert.ToBoolean(RHS); },
            [Tokens.XOr] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) ^ Convert.ToBoolean(RHS); },
        };

        private static Dictionary<string, Func<double, double>> MathOpsUnary = new Dictionary<string, Func<double, double>>()
        {
            [MathTokens.ADDITIVE_INVERSE] = delegate (double value) { return value * -1.0; }
        };

        private static Dictionary<string, Func<double, bool>> LogicOpsUnary = new Dictionary<string, Func<double, bool>>()
        {
            [Tokens.Not] = delegate (double value) { return !(Convert.ToBoolean(value)); }
        };

        private static List<string> ResultTypeRanking = new List<string>()
        {
            Tokens.Currency,
            Tokens.Double,
            Tokens.Single,
            Tokens.Long,
            Tokens.Integer,
            Tokens.Byte,
            Tokens.Boolean,
        };

        public UCIValueExpressionEvaluator(IUCIValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IUCIValue Evaluate(IUCIValue LHS, IUCIValue RHS, string opSymbol)
        {
            var isMathOp = MathOpsBinary.ContainsKey(opSymbol);
            var isLogicOp = LogicOpsBinary.ContainsKey(opSymbol);
            Debug.Assert(isMathOp || isLogicOp);

            var opResultTypeName = isMathOp ? DetermineMathResultType(LHS, RHS) : Tokens.Boolean;
            var operands = PrepareOperands(new string[] { LHS.ValueText, RHS.ValueText });

            if (operands.Count == 2)
            {
                if (isMathOp)
                {
                    var mathResult = MathOpsBinary[opSymbol](operands[0], operands[1]);
                    return _valueFactory.Create(mathResult.ToString(), opResultTypeName);
                }
                var logicResult = LogicOpsBinary[opSymbol](operands[0], operands[1]);
                return _valueFactory.Create(logicResult.ToString(), opResultTypeName);
            }
            return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opResultTypeName);
        }

        public IUCIValue Evaluate(IUCIValue value, string opSymbol)
        {
            var isMathOp = MathOpsUnary.ContainsKey(opSymbol);
            var isLogicOp = LogicOpsUnary.ContainsKey(opSymbol);
            Debug.Assert(isMathOp || isLogicOp);

            var opResultTypeName = isMathOp ? value.TypeName : Tokens.Boolean;
            var operands = PrepareOperands(new string[] { value.ValueText });
            if (operands.Count == 1)
            {
                if (isMathOp)
                {
                    var mathResult = MathOpsUnary[opSymbol](operands[0]);
                    return _valueFactory.Create(mathResult.ToString(), opResultTypeName);
                }
                var logicResult = LogicOpsUnary[opSymbol](operands[0]);
                return _valueFactory.Create(logicResult.ToString(), opResultTypeName);
            }
            return _valueFactory.Create($"{opSymbol} {value.ValueText}", opResultTypeName);
        }

        private static string DetermineMathResultType(IUCIValue LHS, IUCIValue RHS)
        {
            var lhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(LHS.TypeName));
            var rhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(RHS.TypeName));
            return lhsTypeNameIndex <= rhsTypeNameIndex ? LHS.TypeName : RHS.TypeName;
        }

        private static List<double> PrepareOperands(string[] args)
        {
            var results = new List<double>();
            foreach (var arg in args)
            {
                string parseArg = arg;
                if (arg.Equals(Tokens.True) || arg.Equals(Tokens.False))
                {
                    parseArg = arg.Equals(Tokens.True) ? "-1" : "0";
                }

                if (double.TryParse(parseArg, out double result))
                {
                    results.Add(result);
                }
            }
            return results;
        }
    }
}
