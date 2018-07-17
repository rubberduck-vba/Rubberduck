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
        IParseTreeValue Evaluate(IParseTreeValue LHS, string opSymbol);
    }

    public class ParseTreeExpressionEvaluator : IParseTreeExpressionEvaluator
    {
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly bool _isOptionCompareBinary;

        public ParseTreeExpressionEvaluator(IParseTreeValueFactory valueFactory, bool isOptionCompareBinary = true)
        {
            _valueFactory = valueFactory;
            _isOptionCompareBinary = isOptionCompareBinary;
        }

        public IParseTreeValue Evaluate(IParseTreeValue LHS, IParseTreeValue RHS, string opSymbol)
        {
            if (!(IsSupportedSymbol(opSymbol)))
            {
                throw new ArgumentException($"Unsupported operation ({opSymbol}) passed to Evaluate function");
            }

            if (opSymbol.Equals(LogicalOperators.NOT))
            {
                throw new ArgumentException($"Unary operator ({opSymbol}) passed to binary Evaluate function");
            }

            if (ArithmeticOperators.Includes(opSymbol))
            {
                return EvaluateArithmeticOp(opSymbol, LHS, RHS);
            }

            if (RelationalOperators.Includes(opSymbol))
            {
                return EvaluateRelationalOp(opSymbol, LHS, RHS);
            }

            return EvaluateLogicalOperator(opSymbol, LHS, RHS);
        }

        public IParseTreeValue Evaluate(IParseTreeValue parseTreeValue, string opSymbol)
        {
            if (!(opSymbol.Equals(ArithmeticOperators.ADDITIVE_INVERSE)
                || opSymbol.Equals(LogicalOperators.NOT)))
            {
                throw new ArgumentException($"Binary operator ({opSymbol}) passed to unary evaluation function");
            }

            if (opSymbol.Equals(ArithmeticOperators.ADDITIVE_INVERSE))
            {
                return EvaluateUnaryMinus(parseTreeValue);
            }
            return EvaluateLogicalNot(parseTreeValue);
        }

        private bool IsStringCompare(IParseTreeValue LHS, IParseTreeValue RHS)
             => (LHS.TypeName == Tokens.String) && (RHS.TypeName == Tokens.String);

        private IParseTreeValue EvaluateRelationalOp(string opSymbol, IParseTreeValue LHS, IParseTreeValue RHS)
        {
            var opProvider = new OperatorDeclaredTypeProvider((LHS.TypeName, RHS.TypeName), opSymbol);
            if (!(LHS.ParsesToConstantValue && RHS.ParsesToConstantValue))
            {
                //special case of resolve-able expression with variable LHS
                if (opSymbol.Equals(Tokens.Like) && RHS.ValueText.Equals("*"))
                {
                    return _valueFactory.Create(Tokens.True, Tokens.Boolean);
                }
                //Unable to resolve to a value, return an expression
                if (opProvider.OperatorDeclaredType.Equals(string.Empty))
                {
                    return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", Tokens.Variant);
                }
                return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(RelationalOperators.EQ))
            {
                var result = IsStringCompare(LHS, RHS) ? 
                            Compare(LHS, RHS, (string a, string b) => { return AreEqual(a,b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a == b; }, (double a, double b) => { return a == b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(RelationalOperators.NEQ))
            {
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return !AreEqual(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a != b; }, (double a, double b) => { return a != b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(RelationalOperators.LT))
            {
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsLessThan(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a < b; }, (double a, double b) => { return a < b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(RelationalOperators.LTE) || opSymbol.Equals(RelationalOperators.LTE2))
            {
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsLessThan(a, b) || AreEqual(a,b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a <= b; }, (double a, double b) => { return a <= b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(RelationalOperators.GT))
            {
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsGreaterThan(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a > b; }, (double a, double b) => { return a > b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(RelationalOperators.GTE) || opSymbol.Equals(RelationalOperators.GTE2))
            {
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsGreaterThan(a, b) || AreEqual(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a >= b; }, (double a, double b) => { return a >= b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(RelationalOperators.LIKE))
            {
                if (RHS.ValueText.Equals("*"))
                {
                    return _valueFactory.Create(Tokens.True, Tokens.Boolean);
                }

                if (LHS.ParsesToConstantValue && RHS.ParsesToConstantValue)
                {
                    var matches = Like(LHS.ValueText, RHS.ValueText);
                    return _valueFactory.Create(matches.ToString(), Tokens.Boolean);
                }
            }
            return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opProvider.OperatorDeclaredType);
        }

        private IParseTreeValue EvaluateLogicalNot(IParseTreeValue parseTreeValue)
        {
            var opProvider = new OperatorDeclaredTypeProvider(parseTreeValue.TypeName, LogicalOperators.NOT);
            if (!parseTreeValue.ParsesToConstantValue)
            {
                //Unable to resolve to a value, return an expression
                var opType = opProvider.OperatorDeclaredType;
                opType = opType.Equals(string.Empty) ? Tokens.Variant : opProvider.OperatorDeclaredType;
                return _valueFactory.Create($"{LogicalOperators.NOT} {parseTreeValue.ValueText}", opType);
            }

            if (parseTreeValue.TryConvertValue(out long value))
            {
                return _valueFactory.Create((~value).ToString(CultureInfo.InvariantCulture), opProvider.OperatorDeclaredType);
            }
            throw new OverflowException($"Unable to convert {parseTreeValue} to Long");
        }

        private IParseTreeValue EvaluateLogicalOperator(string opSymbol, IParseTreeValue LHS, IParseTreeValue RHS)
        {
            var opProvider = new OperatorDeclaredTypeProvider((LHS.TypeName, RHS.TypeName), opSymbol);
            if (!(LHS.ParsesToConstantValue && RHS.ParsesToConstantValue))
            {
                //Unable to resolve to a value, return an expression
                var opType = opProvider.OperatorDeclaredType;
                opType = opType.Equals(string.Empty) ? Tokens.Variant : opType;
                return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opType);
            }

            if (!(OperatorDeclaredTypeProvider.IntegralNumericTypes.Contains(opProvider.OperatorDeclaredType)))
            {
                return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(LogicalOperators.AND))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return a && b; })
                    : Calculate(LHS, RHS, (long a, long b) => { return a & b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(LogicalOperators.OR))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return a || b; })
                    : Calculate(LHS, RHS, (long a, long b) => { return a | b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(LogicalOperators.XOR))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return a ^ b; })
                    : Calculate(LHS, RHS, (long a, long b) => { return a ^ b; });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(LogicalOperators.EQV))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return Eqv(a, b); })
                    : Calculate(LHS, RHS, (long a, long b) => { return Eqv(a, b); });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(LogicalOperators.IMP))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return Imp(a, b); })
                    : Calculate(LHS, RHS, (long a, long b) => { return Imp(a, b); });
                return _valueFactory.Create(result, opProvider.OperatorDeclaredType);
            }

            return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opProvider.OperatorDeclaredType);
        }

        private IParseTreeValue EvaluateUnaryMinus(IParseTreeValue parseTreeValue)
        {
            var opProvider = new OperatorDeclaredTypeProvider(parseTreeValue.TypeName, ArithmeticOperators.ADDITIVE_INVERSE);
            if (!parseTreeValue.ParsesToConstantValue)
            {
                //Unable to resolve to a value, return an expression
                var opTypeName = opProvider.OperatorDeclaredType;
                opTypeName = opTypeName.Equals(string.Empty) ? Tokens.Variant : opTypeName;
                return _valueFactory.Create($"{ArithmeticOperators.ADDITIVE_INVERSE} {parseTreeValue.ValueText}", opTypeName);
            }

            var declaredTypeName = opProvider.OperatorDeclaredType;
            if (parseTreeValue.TryConvertValue(out decimal decValue))
            {
                return _valueFactory.Create((0 - decValue).ToString(CultureInfo.InvariantCulture), declaredTypeName);
            }
            if (parseTreeValue.TryConvertValue(out double dblValue))
            {
                return _valueFactory.Create((0 - dblValue).ToString(CultureInfo.InvariantCulture), declaredTypeName);
            }
            throw new ArgumentException($"Unable to process opSymbol: {ArithmeticOperators.ADDITIVE_INVERSE}");
        }

        private IParseTreeValue EvaluateArithmeticOp(string opSymbol, IParseTreeValue LHS, IParseTreeValue RHS)
        {
            Debug.Assert(ArithmeticOperators.Includes(opSymbol));

            var opProvider = new OperatorDeclaredTypeProvider((LHS.TypeName,RHS.TypeName), opSymbol);
            if (!(LHS.ParsesToConstantValue && RHS.ParsesToConstantValue))
            {
                //Unable to resolve to a value, return an expression
                var opTypeName = opProvider.OperatorDeclaredType;
                opTypeName = opTypeName.Equals(string.Empty) ? Tokens.Variant : opTypeName;
                return _valueFactory.Create($"{LHS.ValueText} {opSymbol} {RHS.ValueText}", opTypeName);
            }

            if (opSymbol.Equals(ArithmeticOperators.MULTIPLY))
            {
                return _valueFactory.Create(Calculate(LHS, RHS, (decimal a, decimal b) => { return a * b; }, (double a, double b) => { return a * b; }), opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(ArithmeticOperators.DIVIDE))
            {
                return _valueFactory.Create(Calculate(LHS, RHS, (decimal a, decimal b) => { return a / b; }, (double a, double b) => { return a / b; }), opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(ArithmeticOperators.INTEGER_DIVIDE))
            {
                return _valueFactory.Create(Calculate(LHS, RHS, IntDivision, IntDivision), opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(ArithmeticOperators.PLUS))
            {
                if (LHS.TypeName.Equals(RHS.TypeName) && LHS.TypeName.Equals(Tokens.String))
                {
                    return _valueFactory.Create($"{Concat(LHS.ValueText, RHS.ValueText)}", Tokens.String);
                }
                return _valueFactory.Create(Calculate(LHS, RHS, (decimal a, decimal b) => { return a + b; }, (double a, double b) => { return a + b; }), opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(ArithmeticOperators.MINUS))
            {
                //TODO: Add exception case when Date type is supported (Date - Date => Double)
                return _valueFactory.Create(Calculate(LHS, RHS, (decimal a, decimal b) => { return a - b; }, (double a, double b) => { return a - b; }), opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(ArithmeticOperators.EXPONENT))
            {
                //Exponent always results in a Double
                return _valueFactory.Create(Calculate(LHS, RHS, null, (double a, double b) => { return Math.Pow(a, b); }), opProvider.OperatorDeclaredType);
            }
            else if (opSymbol.Equals(ArithmeticOperators.MODULO))
            {
                return _valueFactory.Create(Calculate(LHS, RHS, (decimal a, decimal b) => { return a % b; }, (double a, double b) => { return a % b; }), opProvider.OperatorDeclaredType);
            }

            //ArithmeticOperators.AMPERSAND
            return _valueFactory.Create($"{Concat(LHS.ValueText, RHS.ValueText)}", Tokens.String);
        }

        private decimal IntDivision(decimal lhs, decimal rhs) => Math.Truncate(Convert.ToDecimal(Convert.ToInt64(lhs) / Convert.ToInt64(rhs)));

        private double IntDivision(double lhs, double rhs) => Math.Truncate(Convert.ToDouble(Convert.ToInt64(lhs) / Convert.ToInt64(rhs)));

        private string Calculate(IParseTreeValue LHS, IParseTreeValue RHS, Func<decimal, decimal, decimal> DecCalc, Func<double, double, double> DblCalc)
        {
            if (!(DecCalc is null) && LHS.TryConvertValue(out decimal lhsValue) && RHS.TryConvertValue(out decimal rhsValue))
            {
                return DecCalc(lhsValue, rhsValue).ToString();
            }
            else if (!(DblCalc is null) && LHS.TryConvertValue(out double lhsDblValue) && RHS.TryConvertValue(out double rhsDblValue))
            {
                return DblCalc(lhsDblValue, rhsDblValue).ToString();
            }
            throw new OverflowException();
        }

        private string Compare(IParseTreeValue LHS, IParseTreeValue RHS, Func<decimal, decimal, bool> DecCalc, Func<double, double, bool> DblCalc)
        {
            if (!(DecCalc is null) && LHS.TryConvertValue(out decimal lhsValue) && RHS.TryConvertValue(out decimal rhsValue))
            {
                return DecCalc(lhsValue, rhsValue) ? Tokens.True : Tokens.False;
            }
            else if (!(DblCalc is null) && LHS.TryConvertValue(out double lhsDblValue) && RHS.TryConvertValue(out double rhsDblValue))
            {
                return DblCalc(lhsDblValue, rhsDblValue) ? Tokens.True : Tokens.False;
            }
            throw new OverflowException();
        }

        private string Compare(IParseTreeValue LHS, IParseTreeValue RHS, Func<string, string, bool> StringComp)
        {
            if (!(StringComp is null))
            {
                return StringComp(LHS.ValueText, RHS.ValueText) ? Tokens.True : Tokens.False;
            }
            throw new ArgumentNullException();
        }

        private string Calculate(IParseTreeValue LHS, IParseTreeValue RHS, Func<long, long, long> LogicCalc)
        {
            if (!(LogicCalc is null) && LHS.TryConvertValue(out long lhsValue) && RHS.TryConvertValue(out long rhsValue))
            {
                return LogicCalc(lhsValue, rhsValue).ToString();
            }
            throw new ArgumentNullException();
        }

        private string Calculate(IParseTreeValue LHS, IParseTreeValue RHS, Func<bool, bool, bool> LogicCalc)
        {
            if (!(LogicCalc is null) && LHS.TryConvertValue(out long lhsValue) && RHS.TryConvertValue(out long rhsValue))
            {
                return LogicCalc(lhsValue != 0, rhsValue != 0).ToString();
            }
            throw new ArgumentNullException();
        }

        private bool IsSupportedSymbol(string opSymbol)
        {
            return ArithmeticOperators.Includes(opSymbol)
                || RelationalOperators.Includes(opSymbol)
                || LogicalOperators.Incudes(opSymbol);
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
            var compareOptions = _isOptionCompareBinary ? 
                StringComparison.CurrentCulture | StringComparison.Ordinal 
                : StringComparison.CurrentCulture | StringComparison.OrdinalIgnoreCase;
            return String.Equals(lhs, rhs, compareOptions);
        }

        private bool IsLessThan(string lhs, string rhs)
        {
            var compareOptions = _isOptionCompareBinary ? CompareOptions.None : CompareOptions.IgnoreCase;
            return String.Compare(lhs, rhs, CultureInfo.CurrentCulture, compareOptions) < 0;
        }

        private bool IsGreaterThan(string lhs, string rhs)
        {
            var compareOptions = _isOptionCompareBinary ? CompareOptions.None : CompareOptions.IgnoreCase;
            return String.Compare(lhs, rhs, CultureInfo.CurrentCulture, compareOptions) > 0;
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

    public class ArithmeticOperators
    {
        private static string _multiply;
        private static string _divide;
        private static string _plus;
        private static string _minusSign;
        private static string _exponent;
        private static string _ampersand;
        private static string _integerDivide;

        public static string MULTIPLY => _multiply ?? LoadSymbols(VBAParser.MULT);
        public static string DIVIDE => _divide ?? LoadSymbols(VBAParser.DIV);
        public static string INTEGER_DIVIDE => _integerDivide ?? LoadSymbols(VBAParser.INTDIV);
        public static string PLUS => _plus ?? LoadSymbols(VBAParser.PLUS);
        public static string MINUS => _minusSign ?? LoadSymbols(VBAParser.MINUS);
        public static string ADDITIVE_INVERSE => MINUS;
        public static string EXPONENT => _exponent ?? LoadSymbols(VBAParser.POW);
        public static string MODULO => Tokens.Mod;
        public static string AMPERSAND => _ampersand ?? LoadSymbols(VBAParser.AMPERSAND);

        public static bool Includes(string opSymbol) => SymbolList.Contains(opSymbol);

        public static List<string> SymbolList = new List<string>()
        {
            MULTIPLY,
            DIVIDE,
            INTEGER_DIVIDE,
            PLUS,
            MINUS,
            ADDITIVE_INVERSE,
            EXPONENT,
            MODULO,
            AMPERSAND,
        };

        private static string LoadSymbols(int target)
        {
            _multiply = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.MULT).Replace("'", "");
            _divide = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.DIV).Replace("'", "");
            _integerDivide = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.INTDIV).Replace("'", "");
            _plus = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.PLUS).Replace("'", "");
            _minusSign = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.MINUS).Replace("'", "");
            _exponent = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.POW).Replace("'", "");
            _ampersand = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.AMPERSAND).Replace("'", "");
            return VBAParser.DefaultVocabulary.GetLiteralName(target).Replace("'", "");
        }
    }

    public class RelationalOperators
    {
        private static string _lessThan;
        private static string _greaterThan;
        private static string _equalTo;

        public static string EQ => _equalTo ?? LoadSymbols(VBAParser.EQ);
        public static string NEQ => "<>";
        public static string LT => _lessThan ?? LoadSymbols(VBAParser.LT);
        public static string LTE => "<=";
        public static string LTE2 => "=<";
        public static string GT => _greaterThan ?? LoadSymbols(VBAParser.GT);
        public static string GTE => ">=";
        public static string GTE2 => "=>";
        public static string LIKE => Tokens.Like;

        public static bool Includes(string opSymbol) => SymbolList.Contains(opSymbol);

        public static List<string> SymbolList = new List<string>()
        {
            EQ,
            NEQ,
            LT,
            LTE,
            LTE2,
            GT,
            GTE,
            GTE2,
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

    public class LogicalOperators
    {
        public static string AND => Tokens.And;
        public static string OR => Tokens.Or;
        public static string XOR => Tokens.XOr;
        public static string NOT => Tokens.Not;
        public static string EQV => Tokens.Eqv;
        public static string IMP => Tokens.Imp;

        public static bool Incudes(string opSymbol) => SymbolList.Contains(opSymbol);

        public static List<string> SymbolList = new List<string>()
        {
            AND,
            OR,
            XOR,
            NOT,
            EQV,
            IMP,
        };
    }
}
