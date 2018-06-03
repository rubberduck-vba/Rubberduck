using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

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

        //TODO: Review to get these Dictionaries back to private
        public static readonly Dictionary<string, Func<double, double, double>> MathOpsBinary = new Dictionary<string, Func<double, double, double>>()
        {
            [MathSymbols.MULTIPLY] = delegate (double LHS, double RHS) { return LHS * RHS; },
            [MathSymbols.DIVIDE] = delegate (double LHS, double RHS) { return LHS / RHS; },
            [MathSymbols.INTEGER_DIVIDE] = delegate (double LHS, double RHS) { return Math.Truncate(Convert.ToDouble(Convert.ToInt64(LHS) / Convert.ToInt64(RHS))); },
            [MathSymbols.PLUS] = delegate (double LHS, double RHS) { return LHS + RHS; },
            [MathSymbols.MINUS] = delegate (double LHS, double RHS) { return LHS - RHS; },
            [MathSymbols.EXPONENT] = Math.Pow,
            [MathSymbols.MODULO] = delegate (double LHS, double RHS) { return LHS % RHS; },
        };

        public static readonly Dictionary<string, Func<double, double, bool>> LogicOpsBinary = new Dictionary<string, Func<double, double, bool>>()
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

        public static readonly Dictionary<string, Func<double, double>> MathOpsUnary = new Dictionary<string, Func<double, double>>()
        {
            [MathSymbols.ADDITIVE_INVERSE] = delegate (double value) { return value * -1.0; }
        };

        public static readonly Dictionary<string, Func<double, bool>> LogicOpsUnary = new Dictionary<string, Func<double, bool>>()
        {
            [LogicSymbols.NOT] = delegate (double value) { return !(Convert.ToBoolean(value)); }
        };

        public static Dictionary<string, Func<string, string, bool>> LogicOpsString = new Dictionary<string, Func<string, string, bool>>()
        {
            [Tokens.Like] = VBALogicOperators.Like
        };

        private static readonly List<string> ResultTypeRanking = new List<string>()
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
            var isBinaryStringOp = LogicOpsString.ContainsKey(opSymbol);
            Debug.Assert(IsSupportedSymbol(opSymbol));
            //Debug.Assert(isMathOp || isLogicOp);

            //var opResultTypeName = isMathOp ? DetermineMathResultType(LHS, RHS) : Tokens.Boolean;
            //var operands = PrepareOperands(new IParseTreeValue[] { LHS, RHS });
            var (lhs,rhs) = PrepareOperands(LHS, RHS);

            if (!lhs.typeName.Equals(string.Empty) && !rhs.typeName.Equals(string.Empty))  //operands.Count == 2)
            {
                if (isMathOp)
                {
                    //var mathResult = MathOpsBinary[opSymbol](operands[0], operands[1]);
                    var mathResult = MathOpsBinary[opSymbol](lhs.value, rhs.value);
                    //return _valueFactory.Create(mathResult.ToString(CultureInfo.InvariantCulture), opResultTypeName);
                    return _valueFactory.Create(mathResult.ToString(CultureInfo.InvariantCulture), DetermineMathResultType(lhs.typeName, rhs.typeName));
                }
                else if (isLogicOp)
                {
                    if (opSymbol.Equals(LogicSymbols.EQV) || opSymbol.Equals(LogicSymbols.IMP))
                    {
                        if (LHS.TypeName == RHS.TypeName
                            && (LHS.TypeName.Equals(Tokens.Long) || LHS.TypeName.Equals(Tokens.Integer)))
                        {
                            var lhsOperand = Convert.ToInt64(lhs.value);
                            var rhsOperand = Convert.ToInt64(rhs.value);
                            var result = opSymbol.Equals(LogicSymbols.EQV) ?
                                VBALogicOperators.Eqv(lhsOperand, rhsOperand)
                                : VBALogicOperators.Imp(lhsOperand, rhsOperand);
                            return _valueFactory.Create(result.ToString(), Tokens.Long);
                        }
                    }
                    if (LogicOpsBinary.ContainsKey(opSymbol))
                    {
                        var logicResult = LogicOpsBinary[opSymbol](lhs.value, rhs.value);
                        return _valueFactory.Create(logicResult.ToString(), Tokens.Boolean);
                    }

                }
                //var logicResult = LogicOpsBinary[opSymbol](operands[0], operands[1]);
                //return _valueFactory.Create(logicResult.ToString(), opResultTypeName);
            }

            if (isBinaryStringOp)
            {
                if (LHS.ParsesToConstantValue && RHS.ParsesToConstantValue)
                {
                    var stringOpResult = LogicOpsString[opSymbol](LHS.ValueText, RHS.ValueText);
                    return _valueFactory.Create(stringOpResult.ToString(), Tokens.Boolean);
                }
            }
            var opResultTypeName = isMathOp ? DetermineMathResultType(LHS.TypeName, RHS.TypeName) : Tokens.Boolean;
            return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opResultTypeName);
            //return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}",
            //    isMathOp ? DetermineMathResultType(LHS.TypeName, RHS.TypeName) : Tokens.Boolean); //opResultTypeName);
        }

        public IParseTreeValue Evaluate(IParseTreeValue value, string opSymbol, string requestedResultType)
        {
            var isMathOp = MathOpsUnary.ContainsKey(opSymbol);
            var isLogicOp = LogicOpsUnary.ContainsKey(opSymbol);
            Debug.Assert(isMathOp || isLogicOp);

            //var operands = PrepareOperands(new IParseTreeValue[] { value });
            var operand = PrepareOperand(value);
            if (!operand.value.Equals(string.Empty))  //operands.Count == 1)
            {
                if (isMathOp)
                {
                    //var mathResult = MathOpsUnary[opSymbol](operands[0]);
                    var mathResult = MathOpsUnary[opSymbol](operand.value);
                    return _valueFactory.Create(mathResult.ToString(CultureInfo.InvariantCulture), requestedResultType);
                }

                //Unary Not (!) operator
                if  (!value.TypeName.Equals(Tokens.Boolean) &&  value.TryConvertValue(out long opValue))
                {
                    var bitwiseComplement = ~opValue;
                    return _valueFactory.Create(Convert.ToBoolean(bitwiseComplement).ToString(), requestedResultType);
                }

                if  (value.TypeName.Equals(Tokens.Boolean))
                {
                    var logicResult = LogicOpsUnary[opSymbol](operand.value);
                    //var logicResult = LogicOpsUnary[opSymbol](operands[0]);
                    return _valueFactory.Create(logicResult.ToString(), requestedResultType);
                }
            }
            return _valueFactory.Create($"{opSymbol} {value.ValueText}", requestedResultType);
        }

        //private static string DetermineMathResultType(IParseTreeValue LHS, IParseTreeValue RHS)
        //{
        //    var lhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(LHS.TypeName));
        //    var rhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(RHS.TypeName));
        //    return lhsTypeNameIndex <= rhsTypeNameIndex ? LHS.TypeName : RHS.TypeName;
        //}

        private static string DetermineMathResultType(string lhsTypeName, string rhsTypeName)
        {
            var lhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(lhsTypeName));
            var rhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(rhsTypeName));
            return lhsTypeNameIndex <= rhsTypeNameIndex ? lhsTypeName : rhsTypeName;
        }

        //private static List<double> PrepareOperands(string[] args)
        //{
        //    var results = new List<double>();
        //    foreach (var arg in args)
        //    {
        //        string parseArg = arg;
        //        if (arg.Equals(Tokens.True) || arg.Equals(Tokens.False))
        //        {
        //            parseArg = arg.Equals(Tokens.True) ? "-1" : "0";
        //        }

        //        if (double.TryParse(parseArg, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
        //        {
        //            results.Add(result);
        //        }
        //    }
        //    return results;
        //}

        //private static List<double> PrepareOperands(IParseTreeValue[] args)
        //{
        //    var results = new List<double>();
        //    foreach (var arg in args)
        //    {
        //        if (arg.TryConvertValue(out double value))
        //        {
        //            results.Add(value);
        //        }
        //    }
        //    return results;
        //}

        private static ((string typeName, double value) lhs, (string typeName, double value) rhs)
            PrepareOperands(IParseTreeValue LHS, IParseTreeValue RHS)
        {
            return (PrepareOperand(LHS), PrepareOperand(RHS));
        }

        //<<<<<<< HEAD
        private static (string typeName, double value) PrepareOperand(IParseTreeValue parseTreeValue)
        {
            if (!parseTreeValue.ParsesToConstantValue)
            {
                return (string.Empty, default);
            }
            (string typeName, double value) lhs = (string.Empty, default);
            if (parseTreeValue.TryConvertValue(out double value))
            {
                //results.Add(value);
                lhs = (parseTreeValue.TypeName, value);
            }
/*
            if (parseTreeValue.TypeName.Equals(Tokens.Currency))
            {
                //TODO: does this need to find its way to to different math/logic op?
                //lhs = (Tokens.Currency, double.Parse(parseTreeValue.ValueText));
                //=======
                //if (double.TryParse(parseArg, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                if(double.TryParse(lhs.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                {
                    //results.Add(result);
                    lhs = (Tokens.Currency, result);
                }
                //>>>>>>> rubberduck-vba/next
            }
            else if (parseTreeValue.TypeName.Equals(Tokens.Long))
            {
                lhs = (Tokens.Long, double.Parse(parseTreeValue.ValueText));
            }
            else if (long.TryParse(parseTreeValue.ValueText, out long result))
            {
                lhs = (Tokens.Long, Convert.ToDouble(result));
            }
            else if (parseTreeValue.ValueText.Equals(Tokens.True) || parseTreeValue.ValueText.Equals(Tokens.False))
            {
                result = parseTreeValue.ValueText.Equals(Tokens.True) ? -1 : 0;
                lhs = (Tokens.Long, Convert.ToDouble(result));
            }
            else if (double.TryParse(parseTreeValue.ValueText, out double dResult))
            {
                lhs = (Tokens.Double, dResult);
            }
*/
            return lhs;
        }

        private bool IsSupportedSymbol(string opSymbol)
        {
            return MathOpsBinary.ContainsKey(opSymbol)
                || MathOpsUnary.ContainsKey(opSymbol)
                || LogicOpsBinary.ContainsKey(opSymbol)
                || LogicOpsUnary.ContainsKey(opSymbol)
                || LogicOpsString.ContainsKey(opSymbol);
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

        public static List<string> MathSymbolList = new List<string>()
        {
            MULTIPLY,
            DIVIDE,
            INTEGER_DIVIDE,
            PLUS,
            MINUS,
            ADDITIVE_INVERSE,
            EXPONENT,
            MODULO,
        };

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
