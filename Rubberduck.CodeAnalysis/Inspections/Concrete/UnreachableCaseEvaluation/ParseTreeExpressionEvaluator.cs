using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Refactorings;
using Rubberduck.Refactoring.ParseTreeValue;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal interface IParseTreeExpressionEvaluator
    {
        IParseTreeValue Evaluate(IParseTreeValue LHS, IParseTreeValue RHS, string opSymbol);
        IParseTreeValue Evaluate(IParseTreeValue LHS, string opSymbol);
    }

    internal class ParseTreeExpressionEvaluator : IParseTreeExpressionEvaluator
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

            var opProvider = new OperatorTypesProvider((LHS.ValueType, RHS.ValueType), opSymbol);
            if (opProvider.IsMismatch)
            {
                return _valueFactory.CreateMismatchExpression($"{LHS.Token} {opSymbol} {RHS.Token}", Tokens.Variant);
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
             => (LHS.ValueType == Tokens.String) && (RHS.ValueType == Tokens.String);

        private IParseTreeValue EvaluateRelationalOp(string opSymbol, IParseTreeValue LHS, IParseTreeValue RHS)
        {

            var opProvider = new OperatorTypesProvider((LHS.ValueType, RHS.ValueType), opSymbol);

            if (!(LHS.ParsesToConstantValue && RHS.ParsesToConstantValue))
            {
                //special case of resolve-able expression with variable LHS
                if (opSymbol.Equals(Tokens.Like) && RHS.Token.Equals($"\"*\""))
                {
                    return _valueFactory.Create(true);
                }
                //Unable to resolve to a value, return an expression
                if (opProvider.OperatorDeclaredType.Equals(string.Empty))
                {
                    return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", Tokens.Variant);
                }
                return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(RelationalOperators.EQ))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Boolean))
                {
                    return _valueFactory.Create(Compare(LHS, RHS, (bool a, bool b) => { return a == b; }));
                }
                var result = IsStringCompare(LHS, RHS) ? 
                            Compare(LHS, RHS, (string a, string b) => { return AreEqual(a,b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a == b; }, (double a, double b) => { return a == b; });
                return _valueFactory.Create(result.Equals(Tokens.True));
            }

            if (opSymbol.Equals(RelationalOperators.NEQ))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Boolean))
                {
                    return _valueFactory.Create(Compare(LHS, RHS, (bool a, bool b) => { return a == true && b == false; }));
                }
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return !AreEqual(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a != b; }, (double a, double b) => { return a != b; });
                return _valueFactory.Create(result.Equals(Tokens.True));
            }

            if (opSymbol.Equals(RelationalOperators.LT))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Boolean))
                {
                    return _valueFactory.Create(Compare(LHS, RHS, (bool a, bool b) => { return (a == true && b == false); }));
                }
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsLessThan(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a < b; }, (double a, double b) => { return a < b; });
                return _valueFactory.Create(result.Equals(Tokens.True));
            }

            if (opSymbol.Equals(RelationalOperators.LTE) || opSymbol.Equals(RelationalOperators.LTE2))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Boolean))
                {
                    return _valueFactory.Create(Compare(LHS, RHS, (bool a, bool b) => { return (a == true && b == false) || a == b; }));
                }
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsLessThan(a, b) || AreEqual(a,b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a <= b; }, (double a, double b) => { return a <= b; });
                return _valueFactory.Create(result.Equals(Tokens.True));
            }

            if (opSymbol.Equals(RelationalOperators.GT))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Boolean))
                {
                    return _valueFactory.Create(Compare(LHS, RHS, (bool a, bool b) => { return a == false && b == true || a == b; }));
                }
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsGreaterThan(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a > b; }, (double a, double b) => { return a > b; });
                return _valueFactory.Create(result.Equals(Tokens.True));
            }

            if (opSymbol.Equals(RelationalOperators.GTE) || opSymbol.Equals(RelationalOperators.GTE2))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Boolean))
                {
                    return _valueFactory.Create(Compare(LHS, RHS, (bool a, bool b) => { return a == false && b == true || a == b; }));
                }
                var result = IsStringCompare(LHS, RHS) ?
                            Compare(LHS, RHS, (string a, string b) => { return IsGreaterThan(a, b) || AreEqual(a, b); })
                            : Compare(LHS, RHS, (decimal a, decimal b) => { return a >= b; }, (double a, double b) => { return a >= b; });
                return _valueFactory.Create(result.Equals(Tokens.True));
            }

            if (opSymbol.Equals(RelationalOperators.LIKE))
            {
                if (RHS.Token.Equals("*"))
                {
                    return _valueFactory.Create(true);
                }

                if (LHS.ParsesToConstantValue && RHS.ParsesToConstantValue)
                {
                    var matches = Like(LHS.Token, RHS.Token);
                    return _valueFactory.Create(matches);
                }
            }
            return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorDeclaredType);
        }

        private IParseTreeValue EvaluateLogicalNot(IParseTreeValue parseTreeValue)
        {
            var opProvider = new OperatorTypesProvider(parseTreeValue.ValueType, LogicalOperators.NOT);
            if (!parseTreeValue.ParsesToConstantValue)
            {
                //Unable to resolve to a value, return an expression
                var opType = opProvider.OperatorDeclaredType;
                opType = opType.Equals(string.Empty) ? Tokens.Variant : opProvider.OperatorDeclaredType;
                return _valueFactory.CreateExpression($"{LogicalOperators.NOT} {parseTreeValue.Token}", opType);
            }

            var value = parseTreeValue.AsLong();
            return _valueFactory.CreateDeclaredType((~value).ToString(CultureInfo.InvariantCulture), opProvider.OperatorDeclaredType);
        }

        private IParseTreeValue EvaluateLogicalOperator(string opSymbol, IParseTreeValue LHS, IParseTreeValue RHS)
        {
            var opProvider = new OperatorTypesProvider((LHS.ValueType, RHS.ValueType), opSymbol);
            if (!(LHS.ParsesToConstantValue && RHS.ParsesToConstantValue))
            {
                //Unable to resolve to a value, return an expression
                var opType = opProvider.OperatorDeclaredType;
                opType = opType.Equals(string.Empty) ? Tokens.Variant : opType;
                return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opType);
            }

            if (!(OperatorTypesProvider.IntegralNumericTypes.Contains(opProvider.OperatorDeclaredType)))
            {
                return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(LogicalOperators.AND))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return a && b; })
                    : Calculate(LHS, RHS, (long a, long b) => { return a & b; });
                return _valueFactory.CreateDeclaredType(result, opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(LogicalOperators.OR))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return a || b; })
                    : Calculate(LHS, RHS, (long a, long b) => { return a | b; });
                return _valueFactory.CreateDeclaredType(result, opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(LogicalOperators.XOR))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return a ^ b; })
                    : Calculate(LHS, RHS, (long a, long b) => { return a ^ b; });
                return _valueFactory.CreateDeclaredType(result, opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(LogicalOperators.EQV))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return Eqv(a, b); })
                    : Calculate(LHS, RHS, (long a, long b) => { return Eqv(a, b); });
                return _valueFactory.CreateDeclaredType(result, opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(LogicalOperators.IMP))
            {
                var result = opProvider.OperatorDeclaredType.Equals(Tokens.Boolean) ?
                    Calculate(LHS, RHS, (bool a, bool b) => { return Imp(a, b); })
                    : Calculate(LHS, RHS, (long a, long b) => { return Imp(a, b); });
                return _valueFactory.CreateDeclaredType(result, opProvider.OperatorDeclaredType);
            }

            return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorDeclaredType);
        }

        private IParseTreeValue EvaluateUnaryMinus(IParseTreeValue parseTreeValue)
        {
            var opProvider = new OperatorTypesProvider(parseTreeValue.ValueType, ArithmeticOperators.ADDITIVE_INVERSE);
            if (!parseTreeValue.ParsesToConstantValue)
            {
                //Unable to resolve to a value, return an expression
                var opTypeName = opProvider.OperatorDeclaredType;
                return _valueFactory.CreateDeclaredType($"{ArithmeticOperators.ADDITIVE_INVERSE} {parseTreeValue.Token}", opTypeName);
            }

            var effTypeName = opProvider.OperatorEffectiveType;
            if (effTypeName.Equals(Tokens.Date))
            {
                var result = DateTime.FromOADate(0 - parseTreeValue.AsDouble());
                var date = new DateValue(result);
                return _valueFactory.CreateDeclaredType(date.AsDate.ToString(CultureInfo.InvariantCulture), effTypeName);
            }

            var declaredTypeName = opProvider.OperatorDeclaredType;
            if (parseTreeValue.TryLetCoerce(out decimal decValue))
            {
                return _valueFactory.CreateDeclaredType((0 - decValue).ToString(CultureInfo.InvariantCulture), declaredTypeName);
            }
            return _valueFactory.CreateDeclaredType((0 - parseTreeValue.AsDouble()).ToString(CultureInfo.InvariantCulture), declaredTypeName);
        }

        private IParseTreeValue EvaluateArithmeticOp(string opSymbol, IParseTreeValue LHS, IParseTreeValue RHS)
        {
            Debug.Assert(ArithmeticOperators.Includes(opSymbol));

            var opProvider = new OperatorTypesProvider((LHS.ValueType, RHS.ValueType), opSymbol);
            if (!(LHS.ParsesToConstantValue && RHS.ParsesToConstantValue))
            {
                //Unable to resolve to a value, return an expression
                return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorDeclaredType);
            }

            if (!LHS.TryLetCoerce(opProvider.OperatorEffectiveType, out var effLHS)
                || !RHS.TryLetCoerce(opProvider.OperatorEffectiveType, out var effRHS))
            {
                return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorDeclaredType);
            }

            if (opProvider.OperatorEffectiveType.Equals(Tokens.Date))
            {
                if (!(LHS.TryLetCoerce(Tokens.Double, out effLHS) && RHS.TryLetCoerce(Tokens.Double, out effRHS)))
                {
                    return _valueFactory.CreateExpression($"{LHS.Token} {opSymbol} {RHS.Token}", opProvider.OperatorEffectiveType);
                }
            }

            if (opSymbol.Equals(ArithmeticOperators.MULTIPLY))
            {
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, (decimal a, decimal b) => a * b, (double a, double b) => a * b), opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(ArithmeticOperators.DIVIDE))
            {
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, (decimal a, decimal b) => a / b, (double a, double b) => a / b), opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(ArithmeticOperators.INTEGER_DIVIDE))
            {
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, IntDivision, IntDivision), opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(ArithmeticOperators.PLUS))
            {
                if (opProvider.OperatorEffectiveType.Equals(Tokens.String))
                {
                    return _valueFactory.CreateValueType(Concatenate(LHS, RHS), opProvider.OperatorDeclaredType);
                }
                if (opProvider.OperatorEffectiveType.Equals(Tokens.Date))
                {
                    var result = _valueFactory.CreateDeclaredType(Calculate(effLHS, effRHS, null, (double a, double b) => a + b), Tokens.Double);
                    return _valueFactory.CreateDate(result.AsDouble());
                }
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, (decimal a, decimal b) => a + b, (double a, double b) => a + b), opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(ArithmeticOperators.MINUS))
            {
                if (LHS.ValueType.Equals(Tokens.Date) && RHS.ValueType.Equals(Tokens.Date))
                {
                    return _valueFactory.CreateDate(LHS.AsDouble() - RHS.AsDouble());
                }
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, (decimal a, decimal b) => a - b, (double a, double b) => a - b), opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(ArithmeticOperators.EXPONENT))
            {
                //Math.Pow only takes doubles, so the decimal conversion option is null
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, null, Math.Pow), opProvider.OperatorDeclaredType);
            }

            if (opSymbol.Equals(ArithmeticOperators.MODULO))
            {
                return _valueFactory.CreateValueType(Calculate(effLHS, effRHS, (decimal a, decimal b) => a % b, (double a, double b) => a % b), opProvider.OperatorDeclaredType);
            }

            //ArithmeticOperators.AMPERSAND
            return _valueFactory.CreateValueType(Concatenate(LHS, RHS), opProvider.OperatorDeclaredType);
        }

        private string Concatenate(IParseTreeValue LHS, IParseTreeValue RHS)
        {
            var lhs = StripDoubleQuotes(LHS.Token);
            var rhs = StripDoubleQuotes(RHS.Token);
            return $"{ @""""}{lhs}{rhs}{ @""""}";
        }

        private static string StripDoubleQuotes(string input)
        {
            if (input.StartsWith("\""))
            {
                input = input.Substring(1);
            }
            if (input.EndsWith("\""))
            {
                input = input.Substring(0,input.Length - 1);
            }
            return input;
        }

        private decimal IntDivision(decimal lhs, decimal rhs) => Math.Truncate(Convert.ToDecimal(Convert.ToInt64(lhs) / Convert.ToInt64(rhs)));

        private double IntDivision(double lhs, double rhs) => Math.Truncate(Convert.ToDouble(Convert.ToInt64(lhs) / Convert.ToInt64(rhs)));

        private string Calculate(IParseTreeValue LHS, IParseTreeValue RHS, Func<decimal, decimal, decimal> DecimalCalc, Func<double, double, double> DoubleCalc)
        {
            if (DecimalCalc is null && DoubleCalc is null)
            {
                throw new ArgumentNullException();
            }

            if (!(DecimalCalc is null) && LHS.TryLetCoerce(out decimal lhsValue) && RHS.TryLetCoerce(out decimal rhsValue))
            {
                return DecimalCalc(lhsValue, rhsValue).ToString(CultureInfo.InvariantCulture);
            }
            return DoubleCalc(LHS.AsDouble(), RHS.AsDouble()).ToString(CultureInfo.InvariantCulture);
        }

        private string Compare(IParseTreeValue LHS, IParseTreeValue RHS, Func<decimal, decimal, bool> DecimalCompare, Func<double, double, bool> DoubleCompare)
        {
            if (DecimalCompare is null && DoubleCompare is null)
            {
                throw new ArgumentNullException();
            }

            if (!(DecimalCompare is null) && LHS.TryLetCoerce(out decimal lhsValue) && RHS.TryLetCoerce(out decimal rhsValue))
            {
                return DecimalCompare(lhsValue, rhsValue) ? Tokens.True : Tokens.False;
            }
            return DoubleCompare(LHS.AsDouble(), RHS.AsDouble()) ? Tokens.True : Tokens.False;
        }

        private string Compare(IParseTreeValue LHS, IParseTreeValue RHS, Func<string, string, bool> StringComp)
        {
            var compare = StringComp ?? throw new ArgumentNullException();
            return compare(LHS.Token, RHS.Token) ? Tokens.True : Tokens.False;
        }

        private bool Compare(IParseTreeValue LHS, IParseTreeValue RHS, Func<bool, bool, bool> BoolCompare)
        {
            var compare = BoolCompare ?? throw new ArgumentNullException();
            return compare(LHS.AsBoolean(), RHS.AsBoolean());
        }

        private static string Calculate(IParseTreeValue LHS, IParseTreeValue RHS, Func<long, long, long> LogicCalc)
        {
            var calc = LogicCalc ?? throw new ArgumentNullException();
            return calc(LHS.AsLong(), RHS.AsLong()).ToString();
        }

        private static string Calculate(IParseTreeValue LHS, IParseTreeValue RHS, Func<bool, bool, bool> LogicCalc)
        {
            var calc = LogicCalc ?? throw new ArgumentNullException();
            return calc(LHS.AsLong() != 0, RHS.AsLong() != 0).ToString();
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

    internal class ArithmeticOperators
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

    internal class RelationalOperators
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

    internal class LogicalOperators
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
