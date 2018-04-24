using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeExpressionEvaluator
    {
        IParseTreeValue Evaluate(IParseTreeValue LHS, IParseTreeValue RHS, string opSymbol);
        IParseTreeValue Evaluate(IParseTreeValue LHS, string opSymbol, string requestedResultType);
    }

    public class ParseTreeExpressionEvaluator : IParseTreeExpressionEvaluator
    {
        private readonly IParseTreeValueFactory _valueFactory;

        private static Dictionary<string, Func<double, double, double>> MathOpsBinary = new Dictionary<string, Func<double, double, double>>()
        {
            [MathSymbols.MULTIPLY] = delegate (double LHS, double RHS) { return LHS * RHS; },
            [MathSymbols.DIVIDE] = delegate (double LHS, double RHS) { return LHS / RHS; },
            [MathSymbols.INTEGER_DIVIDE] = delegate (double LHS, double RHS) { return Math.Truncate(Convert.ToDouble(Convert.ToInt64(LHS) / Convert.ToInt64(RHS))); },
            [MathSymbols.PLUS] = delegate (double LHS, double RHS) { return LHS + RHS; },
            [MathSymbols.MINUS] = delegate (double LHS, double RHS) { return LHS - RHS; },
            [MathSymbols.EXPONENT] = Math.Pow,
            [MathSymbols.MODULO] = delegate (double LHS, double RHS) { return LHS % RHS; },
        };

        private static Dictionary<string, Func<double, double, bool>> LogicOpsBinary = new Dictionary<string, Func<double, double, bool>>()
        {
            [LogicSymbols.EQ] = delegate (double LHS, double RHS) { return LHS == RHS; },
            [LogicSymbols.NEQ] = delegate (double LHS, double RHS) { return LHS != RHS; },
            [LogicSymbols.LT] = delegate (double LHS, double RHS) { return LHS < RHS; },
            [LogicSymbols.LTE] = delegate (double LHS, double RHS) { return LHS <= RHS; },
            [LogicSymbols.GT] = delegate (double LHS, double RHS) { return LHS > RHS; },
            [LogicSymbols.GTE] = delegate (double LHS, double RHS) { return LHS >= RHS; },
            [LogicSymbols.AND] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) && Convert.ToBoolean(RHS); },
            [LogicSymbols.OR] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) || Convert.ToBoolean(RHS); },
            [LogicSymbols.XOR] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) ^ Convert.ToBoolean(RHS); },
            [LogicSymbols.EQV] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS).Equals(Convert.ToBoolean(RHS)); },
            [LogicSymbols.IMP] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS).Equals(Convert.ToBoolean(RHS)) || Convert.ToBoolean(RHS); },
        };

        private static Dictionary<string, Func<double, double>> MathOpsUnary = new Dictionary<string, Func<double, double>>()
        {
            [MathSymbols.ADDITIVE_INVERSE] = delegate (double value) { return value * -1.0; }
        };

        private static Dictionary<string, Func<double, bool>> LogicOpsUnary = new Dictionary<string, Func<double, bool>>()
        {
            [LogicSymbols.NOT] = delegate (double value) { return !(Convert.ToBoolean(value)); }
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

        public ParseTreeExpressionEvaluator(IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public IParseTreeValue Evaluate(IParseTreeValue LHS, IParseTreeValue RHS, string opSymbol)
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

        public IParseTreeValue Evaluate(IParseTreeValue value, string opSymbol, string requestedResultType)
        {
            var isMathOp = MathOpsUnary.ContainsKey(opSymbol);
            var isLogicOp = LogicOpsUnary.ContainsKey(opSymbol);
            Debug.Assert(isMathOp || isLogicOp);

            var operands = PrepareOperands(new string[] { value.ValueText });
            if (operands.Count == 1)
            {
                if (isMathOp)
                {
                    var mathResult = MathOpsUnary[opSymbol](operands[0]);
                    return _valueFactory.Create(mathResult.ToString(), requestedResultType);
                }

                //Unary Not (!) operator
                if  (!value.TypeName.Equals(Tokens.Boolean) &&  ParseTreeValue.TryConvertValue(value, out long opValue))
                {
                    var bitwiseComplement = ~opValue;
                    return _valueFactory.Create(Convert.ToBoolean(bitwiseComplement).ToString(), requestedResultType);
                }
                else if  (value.TypeName.Equals(Tokens.Boolean))
                {
                    var logicResult = LogicOpsUnary[opSymbol](operands[0]);
                    return _valueFactory.Create(logicResult.ToString(), requestedResultType);
                }
            }
            return _valueFactory.Create($"{opSymbol} {value.ValueText}", requestedResultType);
        }

        private static string DetermineMathResultType(IParseTreeValue LHS, IParseTreeValue RHS)
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

    internal static class MathSymbols
    {
        private static string _multiply;
        private static string _divide;
        private static string _plus;
        private static string _minusSign;
        private static string _exponent;
        private static string _integerDivide;

        public static string MULTIPLY => _multiply ?? LoadSymbols(VBAParser.MULT);
        public static string DIVIDE => _divide ?? LoadSymbols(VBAParser.DIV);
        public static string INTEGER_DIVIDE => _integerDivide ?? LoadSymbols(VBAParser.INTDIV);
        public static string PLUS => _plus ?? LoadSymbols(VBAParser.PLUS);
        public static string MINUS => _minusSign ?? LoadSymbols(VBAParser.MINUS);
        public static string ADDITIVE_INVERSE => MINUS;
        public static string EXPONENT => _exponent ?? LoadSymbols(VBAParser.POW);
        public static string MODULO => Tokens.Mod;

        private static string LoadSymbols(int target)
        {
            _multiply = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.MULT).Replace("'", "");
            _divide = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.DIV).Replace("'", "");
            _integerDivide = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.INTDIV).Replace("'", "");
            _plus = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.PLUS).Replace("'", "");
            _minusSign = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.MINUS).Replace("'", "");
            _exponent = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.POW).Replace("'", "");
            return VBAParser.DefaultVocabulary.GetLiteralName(target).Replace("'", "");
        }
    }
}
