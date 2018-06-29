using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;

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
        private readonly string _ampersand;
        private readonly bool _isOptionCompareBinary;

        private static readonly Dictionary<string, Func<double, double, double>> MathOpsBinary = new Dictionary<string, Func<double, double, double>>()
        {
            [MathSymbols.MULTIPLY] = delegate (double LHS, double RHS) { return LHS * RHS; },
            [MathSymbols.DIVIDE] = delegate (double LHS, double RHS) { return LHS / RHS; },
            [MathSymbols.INTEGER_DIVIDE] = delegate (double LHS, double RHS) { return Math.Truncate(Convert.ToDouble(Convert.ToInt64(LHS) / Convert.ToInt64(RHS))); },
            [MathSymbols.PLUS] = delegate (double LHS, double RHS) { return LHS + RHS; },
            [MathSymbols.MINUS] = delegate (double LHS, double RHS) { return LHS - RHS; },
            [MathSymbols.EXPONENT] = Math.Pow,
            [MathSymbols.MODULO] = delegate (double LHS, double RHS) { return LHS % RHS; },
            [MathSymbols.EQV] = delegate (double LHS, double RHS) { return Eqv(Convert.ToInt64(LHS), Convert.ToInt64(RHS)); },
            [MathSymbols.IMP] = delegate (double LHS, double RHS) { return Imp(Convert.ToInt64(LHS), Convert.ToInt64(RHS)); },
        };

        private static readonly Dictionary<string, Func<double, double, bool>> LogicOpsBinary = new Dictionary<string, Func<double, double, bool>>()
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
        };

        private readonly Dictionary<string, Func<string, string, bool>> LogicOpsString;

        private static readonly Dictionary<string, Func<double, double>> MathOpsUnary = new Dictionary<string, Func<double, double>>()
        {
            [MathSymbols.ADDITIVE_INVERSE] = delegate (double value) { return value * -1.0; }
        };

        private static readonly Dictionary<string, Func<double, bool>> LogicOpsUnary = new Dictionary<string, Func<double, bool>>()
        {
            [LogicSymbols.NOT] = delegate (double value) { return !(Convert.ToBoolean(value)); }
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
            Tokens.String
        };

        public ParseTreeExpressionEvaluator(IParseTreeValueFactory valueFactory, bool isOptionCompareBinary = true)
        {
            _valueFactory = valueFactory;
            _isOptionCompareBinary = isOptionCompareBinary;
            _ampersand = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.AMPERSAND).Replace("'", "");

            LogicOpsString = new Dictionary<string, Func<string, string, bool>>()
            {
                [LogicSymbols.EQ] = delegate (string LHS, string RHS) { return AreEqual(LHS,RHS); },
                [LogicSymbols.NEQ] = delegate (string LHS, string RHS) { return !AreEqual(LHS, RHS); },
                [LogicSymbols.LT] = delegate (string LHS, string RHS) { return IsLessThan(LHS, RHS); },
                [LogicSymbols.LTE] = delegate (string LHS, string RHS) { return AreEqual(LHS, RHS) || IsLessThan(LHS, RHS); },
                [LogicSymbols.GT] = delegate (string LHS, string RHS) { return IsGreaterThan(LHS, RHS); },
                [LogicSymbols.GTE] = delegate (string LHS, string RHS) { return AreEqual(LHS, RHS) || IsGreaterThan(LHS, RHS); },
            };

        }

        public IParseTreeValue Evaluate(IParseTreeValue LHS, IParseTreeValue RHS, string opSymbol)
        {
            var isMathOp = MathOpsBinary.ContainsKey(opSymbol) && !(LHS.TypeName.Equals(Tokens.String) && RHS.TypeName.Equals(Tokens.String));
            var isLogicOp = LogicOpsBinary.ContainsKey(opSymbol);
            var isBinaryStringOp = opSymbol.Equals(Tokens.Like) || opSymbol.Equals(_ampersand) || (LHS.TypeName.Equals(Tokens.String) && (RHS.TypeName.Equals(Tokens.String)));
            Debug.Assert(IsSupportedSymbol(opSymbol));

            (string lhsTypeName, double lhsValue) = PrepareOperand(LHS);
            (string rhsTypeName, double rhsValue) = PrepareOperand(RHS);

            if (!lhsTypeName.Equals(string.Empty) && !rhsTypeName.Equals(string.Empty))
            {
                if (isMathOp)
                {
                    var mathResult = MathOpsBinary[opSymbol](lhsValue, rhsValue);
                    return _valueFactory.Create(mathResult.ToString(CultureInfo.InvariantCulture), DetermineMathResultType(lhsTypeName, rhsTypeName));
                }
                else if (isLogicOp)
                {
                    if (lhsTypeName.Equals(Tokens.String) && rhsTypeName.Equals(Tokens.String))
                    {
                        if (LogicOpsString.ContainsKey(opSymbol))
                        {
                            var stringLogicResult = LogicOpsString[opSymbol](LHS.ValueText, RHS.ValueText);
                            return _valueFactory.Create(stringLogicResult.ToString(CultureInfo.InvariantCulture), Tokens.Boolean);
                        }
                        //Store invalid string logic op result as variable predicate
                        return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", Tokens.Boolean);
                    }
                    var logicResult = LogicOpsBinary[opSymbol](lhsValue, rhsValue);
                    return _valueFactory.Create(logicResult.ToString(CultureInfo.InvariantCulture), Tokens.Boolean);
                }
            }

            if (isBinaryStringOp)
            {
                if (opSymbol.Equals(_ampersand)|| opSymbol.Equals(MathSymbols.PLUS))
                {
                    var concatResult = $"{Concat(LHS.ValueText, RHS.ValueText)}";
                    return _valueFactory.Create(concatResult, Tokens.String);
                }

                if (opSymbol.Contains(Tokens.Like))
                {
                    if (LHS.ParsesToConstantValue && RHS.ParsesToConstantValue)
                    {
                        var stringOpResult = Like(LHS.ValueText, RHS.ValueText);
                        return _valueFactory.Create(stringOpResult.ToString(), Tokens.Boolean);
                    }
                }
            }
            var opResultTypeName = isMathOp ? DetermineMathResultType(LHS.TypeName, RHS.TypeName) : Tokens.Boolean;
            return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opResultTypeName);
        }

        public IParseTreeValue Evaluate(IParseTreeValue value, string opSymbol, string requestedResultType)
        {
            var isMathOp = MathOpsUnary.ContainsKey(opSymbol);
            var isLogicOp = LogicOpsUnary.ContainsKey(opSymbol);
            Debug.Assert(isMathOp || isLogicOp);

            var operand = PrepareOperand(value);
            if (!operand.value.Equals(string.Empty))
            {
                if (isMathOp)
                {
                    var mathResult = MathOpsUnary[opSymbol](operand.value);
                    return _valueFactory.Create(mathResult.ToString(CultureInfo.InvariantCulture), requestedResultType);
                }

                //Unary Not (!) operator
                if (!value.TypeName.Equals(Tokens.Boolean) && value.TryConvertValue(out long opValue))
                {
                    var bitwiseComplement = ~opValue;
                    return _valueFactory.Create(Convert.ToBoolean(bitwiseComplement).ToString(), requestedResultType);
                }

                if (value.TypeName.Equals(Tokens.Boolean))
                {
                    var logicResult = LogicOpsUnary[opSymbol](operand.value);
                    return _valueFactory.Create(logicResult.ToString(), requestedResultType);
                }
            }
            return _valueFactory.Create($"{opSymbol} {value.ValueText}", requestedResultType);
        }

        private static string DetermineMathResultType(string lhsTypeName, string rhsTypeName)
        {
            var lhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(lhsTypeName));
            var rhsTypeNameIndex = ResultTypeRanking.FindIndex(el => el.Equals(rhsTypeName));
            return lhsTypeNameIndex <= rhsTypeNameIndex ? lhsTypeName : rhsTypeName;
        }

        private static (string typeName, double value) PrepareOperand(IParseTreeValue parseTreeValue)
        {
            if (!parseTreeValue.ParsesToConstantValue)
            {
                return (string.Empty, default);
            }
            (string typeName, double value) lhs = (string.Empty, default);
            if (parseTreeValue.TryConvertValue(out double value))
            {
                lhs = (parseTreeValue.TypeName, value);
            }
            else if (parseTreeValue.TypeName.Equals(Tokens.String))
            {
                lhs = (parseTreeValue.TypeName, default);
            }
            return lhs;
        }

        private bool IsSupportedSymbol(string opSymbol)
        {
            return MathOpsBinary.ContainsKey(opSymbol)
                || MathOpsUnary.ContainsKey(opSymbol)
                || LogicOpsBinary.ContainsKey(opSymbol)
                || LogicOpsUnary.ContainsKey(opSymbol)
                || opSymbol.Equals(Tokens.Like)
                || opSymbol.Equals(_ampersand);
        }

        public static bool Eqv(bool lhs, bool rhs) => !(lhs ^ rhs) || (lhs && rhs);

        public static int Eqv(int lhs, int rhs) => ~(lhs ^ rhs) | (lhs & rhs);

        public static long Eqv(long lhs, long rhs) => ~(lhs ^ rhs) | (lhs & rhs);

        public static bool Imp(bool lhs, bool rhs) => rhs || (!lhs && !rhs);

        public static int Imp(int lhs, int rhs) => rhs | (~lhs & ~rhs);

        public static long Imp(long lhs, long rhs) => rhs | (~lhs & ~rhs);

        public static string Concat<T, U>(T lhs, U rhs) => $"{ @""""}{lhs}{rhs}{ @""""}";

        private bool Like(string input, string pattern)
        {
            if (pattern.Equals("*"))
            {
                return true;
            }

            var regexPattern = ConvertLikePatternToRegexPattern(pattern);

            RegexOptions option = _isOptionCompareBinary ? RegexOptions.None : RegexOptions.IgnoreCase;
            var regex = new Regex(regexPattern, option | RegexOptions.CultureInvariant);

            return regex.IsMatch(input);
        }

        private bool AreEqual(string lhs, string rhs)
        {
            return string.Compare(lhs, rhs, !_isOptionCompareBinary, CultureInfo.InvariantCulture) == 0;
        }

        private bool IsLessThan(string lhs, string rhs)
        {
            return string.Compare(lhs, rhs, !_isOptionCompareBinary, CultureInfo.InvariantCulture) < 0;
        }

        private bool IsGreaterThan(string lhs, string rhs)
        {
            return string.Compare(lhs, rhs, !_isOptionCompareBinary, CultureInfo.InvariantCulture) > 0;
        }

        public static string ConvertLikePatternToRegexPattern(string likePattern)
        {
            //The order of replacements matter

            string regexPattern = likePattern;

            //Escape Regex special characters that are not 'Like' special characters
            foreach (var ch in new char[] { '.', '$', '^', '{', '|', '(', ')', '+' })
            {
                regexPattern = Regex.Replace(regexPattern, $"\\{ch}", $"\\{ch}");
            }

            //If the Like pattern does not end with "*", force the last character to match
            regexPattern = $"^{regexPattern}";
            var rgx = new Regex("\\*$");
            regexPattern = rgx.IsMatch(regexPattern) ? rgx.Replace(regexPattern, "[\\D\\d\\s]*") : $"{regexPattern}$";

            //Replace non-escaped *'s with Regex equivalent
            regexPattern = Regex.Replace(regexPattern, "\\*(?=[^\\]])", "[\\D\\d\\s]*");

            //Replace non-escaped ?'s with Regex equivalent
            regexPattern = Regex.Replace(regexPattern, "\\?(?=[^\\]])", ".");

            //Replace non-escaped #'s with Regex equivalent
            regexPattern = Regex.Replace(regexPattern, "\\#(?=[^\\]])", "\\d");

            //Escape Regex special characters that are also escaped 
            //special characters in the Like expressions
            foreach (var ch in new char[] { '*', '?', '[' })
            {
                regexPattern = Regex.Replace(regexPattern, $"\\[\\{ch}]", $"\\{ch}");
            }

            //Replace escaped special character # with Regex equivalent
            regexPattern = Regex.Replace(regexPattern, "\\[#\\]", "#");

            //Replace character group negation with Regex equivalent
            regexPattern = Regex.Replace(regexPattern, "\\[!", "[^");

            return regexPattern;
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
        public static string EQV => Tokens.Eqv;
        public static string IMP => Tokens.Imp;

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
            EQV,
            IMP,
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

    public static class LogicSymbols
    {
        private static string _lessThan;
        private static string _greaterThan;
        private static string _equalTo;

        public static string EQ => _equalTo ?? LoadSymbols(VBAParser.EQ);
        public static string NEQ => "<>";
        public static string LT => _lessThan ?? LoadSymbols(VBAParser.LT);
        public static string LTE => "<=";
        public static string GT => _greaterThan ?? LoadSymbols(VBAParser.GT);
        public static string GTE => ">=";
        public static string AND => Tokens.And;
        public static string OR => Tokens.Or;
        public static string XOR => Tokens.XOr;
        public static string NOT => Tokens.Not;
        public static string LIKE => Tokens.Like;

        public static List<string> LogicSymbolList = new List<string>()
        {
            EQ,
            NEQ,
            LT,
            LTE,
            GT,
            GTE,
            AND,
            OR,
            XOR,
            NOT,
            LIKE,
        };

        private static string LoadSymbols(int target)
        {
            _lessThan = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.LT).Replace("'", "");
            _greaterThan = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.GT).Replace("'", "");
            _equalTo = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.EQ).Replace("'", "");
            return VBAParser.DefaultVocabulary.GetLiteralName(target).Replace("'", "");
        }
    }
}
