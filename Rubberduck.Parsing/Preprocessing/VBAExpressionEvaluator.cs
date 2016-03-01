using Antlr4.Runtime;
using Rubberduck.Parsing.Date;
using Rubberduck.Parsing.Like;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class VBAExpressionEvaluator
    {
        private SymbolTable _symbolTable;
        private readonly Dictionary<Type, ILetCoercion> _letCoercions = new Dictionary<Type, ILetCoercion>()
            {
                { typeof(bool), new BoolLetCoercion() },
                { typeof(byte), new ByteLetCoercion() },
                { typeof(decimal), new DecimalLetCoercion() },
                { typeof(DateTime), new DateLetCoercion() },
                { typeof(string), new StringLetCoercion() },
                { typeof(VBAEmptyValue), new EmptyLetCoercion() }
            };

        public VBAExpressionEvaluator()
            : this(new SymbolTable())
        {
        }

        public VBAExpressionEvaluator(SymbolTable symbolTable)
        {
            _symbolTable = symbolTable;
        }

        public void AddConstant(string name, object value)
        {
            _symbolTable.Add(name, value);
        }

        public void EvaluateConstant(string name, VBAConditionalCompilationParser.CcExpressionContext expression)
        {
            AddConstant(name, Evaluate(expression));
        }

        public bool EvaluateCondition(VBAConditionalCompilationParser.CcExpressionContext expression)
        {
            object condition = Evaluate(expression);
            if (condition == null)
            {
                return false;
            }
            return _letCoercions[condition.GetType()].ToBool(condition);
        }

        private object Evaluate(VBAConditionalCompilationParser.CcExpressionContext expression)
        {
            return Visit(expression);
        }

        private object Visit(VBAConditionalCompilationParser.NameContext context)
        {
            var identifier = context.GetText();
            return _symbolTable.Get(identifier);
        }

        private object VisitUnaryMinus(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var operand = Visit(context.ccExpression()[0]);
            if (operand == null)
            {
                return null;
            }
            else if (operand is DateTime)
            {
                var value = _letCoercions[typeof(DateTime)].ToDecimal(operand);
                value = -value;
                try
                {
                    return _letCoercions[typeof(decimal)].ToDate(value);
                }
                catch
                {
                    // 5.6.9.3.1: If overflow occurs during the coercion to Date, and the operand has a 
                    // declared type of Variant, the result is the Double value.
                    // We don't care about it being a Variant because if it's not a Variant it won't compile/run.
                    // We catch everything because the only case where the code is valid is that if this is an overflow.
                    return value;
                }
            }
            else
            {
                var value = _letCoercions[operand.GetType()].ToDecimal(operand);
                return -value;
            }
        }

        private object VisitUnaryNot(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var operand = Visit(context.ccExpression()[0]);
            if (operand == null)
            {
                return null;
            }
            else if (operand is bool)
            {
                return !(bool)operand;
            }
            else
            {
                var coerced = _letCoercions[operand.GetType()].ToDecimal(operand);
                return (decimal)~Convert.ToInt64(coerced);
            }
        }

        private object VisitPlus(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            if (left is string)
            {
                return (string)left + (string)right;
            }
            else if (left is DateTime || right is DateTime)
            {
                decimal leftValue = _letCoercions[left.GetType()].ToDecimal(left);
                decimal rightValue = _letCoercions[right.GetType()].ToDecimal(right);
                decimal sum = leftValue + rightValue;
                try
                {
                    return _letCoercions[sum.GetType()].ToDate(sum);
                }
                catch
                {
                    return sum;
                }
            }
            else
            {
                var leftNumber = _letCoercions[left.GetType()].ToDecimal(left);
                var rightNumber = _letCoercions[right.GetType()].ToDecimal(right);
                return leftNumber + rightNumber;
            }
        }

        private object VisitMinus(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            else if (left is DateTime && right is DateTime)
            {
                // 5.6.9.3.3 - Effective value type exception.
                // If left + right are both Date then effective value type is double.
                decimal leftValue = _letCoercions[left.GetType()].ToDecimal(left);
                decimal rightValue = _letCoercions[right.GetType()].ToDecimal(right);
                decimal difference = leftValue - rightValue;
                return difference;
            }
            else if (left is DateTime || right is DateTime)
            {
                decimal leftValue = _letCoercions[left.GetType()].ToDecimal(left);
                decimal rightValue = _letCoercions[right.GetType()].ToDecimal(right);
                decimal difference = leftValue - rightValue;
                try
                {
                    return _letCoercions[typeof(decimal)].ToDate(difference);
                }
                catch
                {
                    return difference;
                }
            }
            else
            {
                var leftNumber = _letCoercions[left.GetType()].ToDecimal(left);
                var rightNumber = _letCoercions[right.GetType()].ToDecimal(right);
                return leftNumber - rightNumber;
            }
        }

        private object Visit(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            if (context.literal() != null)
            {
                return Visit(context.literal());
            }
            else if (context.name() != null)
            {
                return Visit(context.name());
            }
            else if (context.L_PAREN() != null)
            {
                return Visit(context.ccExpression()[0]);
            }
            else if (context.MINUS() != null && context.ccExpression().Count == 1)
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
            else if (context.MINUS() != null && context.ccExpression().Count == 2)
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
                return VisitIntrinsicFunction(context);
            }
        }

        private object VisitIntrinsicFunction(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var intrinsicFunction = context.intrinsicFunction();
            var name = intrinsicFunction.intrinsicFunctionName().GetText().ToUpper();
            var expr = Visit(intrinsicFunction.ccExpression());
            if (expr == null)
            {
                return null;
            }
            switch (name)
            {
                case "INT":
                case "FIX":
                    return VisitIntFixFunction(expr);
                case "ABS":
                    return VisitAbsFunction(expr);
                case "SGN":
                    return VisitSgnFunction(expr);
                case "LEN":
                    return VisitLenFunction(expr);
                case "LENB":
                    return VisitLenBFunction(expr);
                case "CBOOL":
                    return VisitCBoolFunction(expr);
                case "CBYTE":
                    return VisitCByteFunction(expr);
                case "CCUR":
                    return VisitCCurFunction(expr);
                case "CDATE":
                    return VisitCDateFunction(expr);
                case "CDBL":
                    return VisitCDblFunction(expr);
                case "CINT":
                    return VisitCIntFunction(expr);
                case "CLNG":
                    return VisitCLngFunction(expr);
                case "CLNGLNG":
                    return VisitCLngLngFunction(expr);
                case "CLNGPTR":
                    return VisitLngPtrFunction(expr);
                case "CSNG":
                    return VisitCSngFunction(expr);
                case "CSTR":
                    return VisitCStrFunction(expr);
                case "CVAR":
                    return VisitCVarFunction(expr);
                default:
                    return null;
            }
        }

        private object VisitCVarFunction(object expr)
        {
            return expr;
        }

        private object VisitCStrFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToString(expr);
        }

        private object VisitCSngFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitLngPtrFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitCLngLngFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitCLngFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitCIntFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitCDblFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitCDateFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDate(expr);
        }

        private object VisitCCurFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToDecimal(expr);
        }

        private object VisitCByteFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToByte(expr);
        }

        private object VisitCBoolFunction(object expr)
        {
            return _letCoercions[expr.GetType()].ToBool(expr);
        }

        private object VisitLenBFunction(object expr)
        {
            return (decimal)_letCoercions[expr.GetType()].ToString(expr).Length * sizeof(Char);
        }

        private object VisitLenFunction(object expr)
        {
            return (decimal)_letCoercions[expr.GetType()].ToString(expr).Length;
        }

        private object VisitSgnFunction(object expr)
        {
            return (decimal)Math.Sign(_letCoercions[expr.GetType()].ToDecimal(expr));
        }

        private object VisitAbsFunction(object expr)
        {
            if (expr is DateTime)
            {
                decimal exprValue = _letCoercions[typeof(DateTime)].ToDecimal(expr);
                exprValue = Math.Abs(exprValue);
                try
                {
                    return _letCoercions[typeof(decimal)].ToDate(exprValue);
                }
                catch
                {
                    return exprValue;
                }
            }
            else
            {
                return Math.Abs(_letCoercions[expr.GetType()].ToDecimal(expr));
            }
        }

        private object VisitIntFixFunction(object expr)
        {
            if (expr is decimal)
            {
                return Math.Truncate((decimal)expr);
            }
            else if (expr is string)
            {
                return Math.Truncate(_letCoercions[typeof(string)].ToDecimal(expr));
            }
            else if (expr is DateTime)
            {
                return _letCoercions[typeof(decimal)].ToDate(Math.Truncate(_letCoercions[typeof(DateTime)].ToDecimal(expr)));
            }
            else
            {
                return _letCoercions[expr.GetType()].ToDecimal(expr);
            }
        }

        private object VisitLike(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var expr = Visit(context.ccExpression()[0]);
            var pattern = Visit(context.ccExpression()[1]);
            if (expr == null || pattern == null)
            {
                return null;
            }
            var exprStr = _letCoercions[expr.GetType()].ToString(expr);
            var patternStr = _letCoercions[pattern.GetType()].ToString(pattern);
            var stream = new AntlrInputStream(patternStr);
            var lexer = new VBALikeLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBALikeParser(tokens);
            var likePattern = parser.likePatternString();
            StringBuilder regexStr = new StringBuilder();
            foreach (var element in likePattern.likePatternElement())
            {
                if (element.NORMALCHAR() != null)
                {
                    regexStr.Append(element.NORMALCHAR().GetText());
                }
                else if (element.QUESTIONMARK() != null)
                {
                    regexStr.Append(".");
                }
                else if (element.HASH() != null)
                {
                    regexStr.Append(@"\d");
                }
                else if (element.STAR() != null)
                {
                    regexStr.Append(@".*?");
                }
                else
                {
                    var charlist = element.likePatternCharlist().GetText();
                    var cleaned = charlist.Replace("[!", "[^");
                    regexStr.Append(cleaned);
                }
            }
            // Full string match, e.g. "abcd" should NOT match "a.c"
            var regex = "^" + regexStr.ToString() + "$";
            return Regex.IsMatch(exprStr, regex);
        }

        private object VisitIs(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            return left == null && right == null;
        }

        private object VisitImp(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null && right == null)
            {
                return null;
            }
            else if (left != null && left is bool && right != null && right is bool)
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                var result = (decimal)(~leftNumber | rightNumber);
                return _letCoercions[result.GetType()].ToBool(result);
            }
            else if (left == null && _letCoercions[right.GetType()].ToDecimal(right) == 0)
            {
                return null;
            }
            else if (left == null)
            {
                return right;
            }
            else if (_letCoercions[left.GetType()].ToDecimal(left) == -1)
            {
                return null;
            }
            else if (right == null)
            {
                return (decimal)(~Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left)) | 0);
            }
            else
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                return (decimal)(~leftNumber | rightNumber);
            }
        }

        private object VisitEqv(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            else if (left is bool && right is bool)
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                var result = (decimal)~(leftNumber ^ rightNumber);
                return _letCoercions[result.GetType()].ToBool(result);
            }
            else
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                return (decimal)~(leftNumber ^ rightNumber);
            }
        }

        private object VisitXor(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            else if (left is bool && right is bool)
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                var result = (decimal)(leftNumber ^ rightNumber);
                return _letCoercions[result.GetType()].ToBool(result);
            }
            else
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                return (decimal)(leftNumber ^ rightNumber);
            }
        }

        private object VisitOr(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null && right == null)
            {
                return null;
            }
            else if (left == null)
            {
                return right;
            }
            else if (right == null)
            {
                return left;
            }
            else if (left is bool && right is bool)
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                var result = (decimal)(leftNumber | rightNumber);
                return _letCoercions[result.GetType()].ToBool(result);
            }
            else
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                return (decimal)(leftNumber | rightNumber);
            }
        }

        private object VisitAnd(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left != null && _letCoercions[left.GetType()].ToDecimal(left) == 0 && right == null)
            {
                return 0;
            }
            else if (right != null && _letCoercions[right.GetType()].ToDecimal(right) == 0 && left == null)
            {
                return 0;
            }
            else if (left == null || right == null)
            {
                return null;
            }
            else if (left is bool && right is bool)
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                var result = (decimal)(leftNumber & rightNumber);
                return _letCoercions[result.GetType()].ToBool(result);
            }
            else
            {
                var leftNumber = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
                var rightNumber = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
                return (decimal)(leftNumber & rightNumber);
            }
        }

        private object VisitGeq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            return !EvaluateLt(left, right);
        }

        private object VisitLeq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            return EvaluateLt(left, right) || EvaluateEq(left, right);
        }

        private object VisitGt(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            return !EvaluateLt(left, right) && !EvaluateEq(left, right);
        }

        private object VisitLt(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            return EvaluateLt(left, right);
        }

        private object VisitNeq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            return !EvaluateEq(left, right);
        }

        private object VisitEq(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            return EvaluateEq(left, right);
        }

        private bool EvaluateLt(object left, object right)
        {
            if (left is string && right is string)
            {
                var leftValue = _letCoercions[left.GetType()].ToString(left);
                var rightValue = _letCoercions[right.GetType()].ToString(right);
                return string.CompareOrdinal(leftValue, rightValue) < 0;
            }
            else if (left is string && right is VBAEmptyValue)
            {
                return string.CompareOrdinal((string)left, _letCoercions[right.GetType()].ToString(right)) < 0;
            }
            else if (right is string && left is VBAEmptyValue)
            {
                return string.CompareOrdinal((string)right, _letCoercions[left.GetType()].ToString(left)) < 0;
            }
            else
            {
                var leftValue = _letCoercions[left.GetType()].ToDecimal(left);
                var rightValue = _letCoercions[right.GetType()].ToDecimal(right);
                return leftValue < rightValue;
            }
        }

        private bool EvaluateEq(object left, object right)
        {
            if (left is string && right is string)
            {
                var leftValue = _letCoercions[left.GetType()].ToString(left);
                var rightValue = _letCoercions[right.GetType()].ToString(right);
                return leftValue == rightValue;
            }
            else if (left is string && right is VBAEmptyValue)
            {
                return (string)left == _letCoercions[right.GetType()].ToString(right);
            }
            else if (right is string && left is VBAEmptyValue)
            {
                return (string)right == _letCoercions[left.GetType()].ToString(left);
            }
            else
            {
                var leftValue = _letCoercions[left.GetType()].ToDecimal(left);
                var rightValue = _letCoercions[right.GetType()].ToDecimal(right);
                return leftValue == rightValue;
            }
        }

        private object VisitConcat(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null && right == null)
            {
                return null;
            }
            string leftValue = string.Empty;
            if (left != null)
            {
                leftValue = _letCoercions[left.GetType()].ToString(left);
            }
            string rightValue = string.Empty;
            if (right != null)
            {
                rightValue = _letCoercions[right.GetType()].ToString(right);
            }
            return leftValue + rightValue;
        }

        private object VisitPow(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            var leftValue = _letCoercions[left.GetType()].ToDecimal(left);
            var rightValue = _letCoercions[right.GetType()].ToDecimal(right);
            return (decimal)Math.Pow(Convert.ToDouble(leftValue), Convert.ToDouble(rightValue));
        }

        private object VisitMod(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            var leftValue = _letCoercions[left.GetType()].ToDecimal(left);
            var rightValue = _letCoercions[right.GetType()].ToDecimal(right);
            return leftValue % rightValue;
        }

        private object VisitIntDiv(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            var leftValue = Convert.ToInt64(_letCoercions[left.GetType()].ToDecimal(left));
            var rightValue = Convert.ToInt64(_letCoercions[right.GetType()].ToDecimal(right));
            return Math.Truncate((decimal)leftValue / rightValue);
        }

        private object VisitMult(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            var leftValue = _letCoercions[left.GetType()].ToDecimal(left);
            var rightValue = _letCoercions[right.GetType()].ToDecimal(right);
            return leftValue * rightValue;
        }

        private object VisitDiv(VBAConditionalCompilationParser.CcExpressionContext context)
        {
            var left = Visit(context.ccExpression()[0]);
            var right = Visit(context.ccExpression()[1]);
            if (left == null || right == null)
            {
                return null;
            }
            var leftValue = _letCoercions[left.GetType()].ToDecimal(left);
            var rightValue = _letCoercions[right.GetType()].ToDecimal(right);
            return leftValue / rightValue;
        }

        private object Visit(VBAConditionalCompilationParser.LiteralContext context)
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
            else if (context.DOUBLELITERAL() != null)
            {
                return VisitDoubleLiteral(context);
            }
            else if (context.INTEGERLITERAL() != null)
            {
                return VisitIntegerLiteral(context);
            }
            else if (context.SHORTLITERAL() != null)
            {
                return VisitShortLiteral(context);
            }
            else if (context.STRINGLITERAL() != null)
            {
                return VisitStringLiteral(context);
            }
            else if (context.TRUE() != null)
            {
                return true;
            }
            else if (context.FALSE() != null)
            {
                return false;
            }
            else if (context.NOTHING() != null || context.NULL() != null)
            {
                return null;
            }
            else if (context.EMPTY() != null)
            {
                return VBAEmptyValue.Value;
            }
            throw new Exception(string.Format("Unexpected literal encountered: {0}", context.GetText()));
        }

        private object VisitStringLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            var str = context.STRINGLITERAL().GetText();
            // Remove quotes
            return str.Substring(1, str.Length - 2);
        }

        private object VisitShortLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            string literal = context.SHORTLITERAL().GetText();
            return decimal.Parse(literal.Replace("#", "").Replace("&", "").Replace("@", ""));
        }

        private object VisitIntegerLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            string literal = context.INTEGERLITERAL().GetText();
            return decimal.Parse(literal.Replace("#", "").Replace("&", "").Replace("@", ""), NumberStyles.Float);
        }

        private object VisitDoubleLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            string literal = context.DOUBLELITERAL().GetText();
            return decimal.Parse(literal.Replace("#", "").Replace("&", "").Replace("@", ""), NumberStyles.Float);
        }

        private object VisitDateLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            string literal = context.DATELITERAL().GetText();
            var stream = new AntlrInputStream(literal);
            var lexer = new VBADateLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBADateParser(tokens);
            var dateLiteral = parser.dateLiteral();
            var dateOrTime = dateLiteral.dateOrTime();
            int year;
            int month;
            int day;
            int hours;
            int mins;
            int seconds;

            Predicate<int> legalMonth = (x) => x >= 0 && x <= 12;
            Func<int, int, int, bool> legalDay = (m, d, y) =>
            {
                bool legalYear = y >= 0 && y <= 32767;
                bool legalM = legalMonth(m);
                bool legalD = false;
                if (legalYear && legalM)
                {
                    int daysInMonth = DateTime.DaysInMonth(y, m);
                    legalD = d >= 1 && d <= daysInMonth;
                }
                return legalYear && legalM && legalD;
            };

            Func<int, int> yearFunc = (x) =>
            {
                if (x >= 0 && x <= 29)
                {
                    return x + 2000;
                }
                else if (x >= 30 && x <= 99)
                {
                    return x + 1900;
                }
                else
                {
                    return x;
                }
            };

            int CY = DateTime.Now.Year;

            if (dateOrTime.dateValue() == null)
            {
                year = 1899;
                month = 12;
                day = 30;
            }
            else
            {
                var dateValue = dateOrTime.dateValue();
                var L = dateValue.dateValuePart()[0];
                var M = dateValue.dateValuePart()[1];
                VBADateParser.DateValuePartContext R = null;
                if (dateValue.dateValuePart().Count == 3)
                {
                    R = dateValue.dateValuePart()[2];
                }
                if (L.DIGIT() != null && M.DIGIT() != null && R == null)
                {
                    var LNumber = int.Parse(L.GetText());
                    var MNumber = int.Parse(M.GetText());
                    if (legalMonth(LNumber) && legalDay(LNumber, MNumber, CY))
                    {
                        month = LNumber;
                        day = MNumber;
                        year = CY;
                    }
                    else if ((legalMonth(MNumber) && legalDay(MNumber, LNumber, CY)))
                    {
                        month = MNumber;
                        day = LNumber;
                        year = CY;
                    }
                    else if (legalMonth(LNumber))
                    {
                        month = LNumber;
                        day = 1;
                        year = MNumber;
                    }
                    else if (legalMonth(MNumber))
                    {
                        month = MNumber;
                        day = 1;
                        year = LNumber;
                    }
                    else
                    {
                        throw new Exception("Invalid date: " + dateLiteral.GetText());
                    }
                }
                else if ((L.DIGIT() != null && M.DIGIT() != null && R != null) && R.DIGIT() != null)
                {
                    var LNumber = int.Parse(L.GetText());
                    var MNumber = int.Parse(M.GetText());
                    var RNumber = int.Parse(R.GetText());
                    if (legalMonth(LNumber) && legalDay(LNumber, MNumber, yearFunc(RNumber)))
                    {
                        month = LNumber;
                        day = MNumber;
                        year = yearFunc(RNumber);
                    }
                    else if (legalMonth(MNumber) && legalDay(MNumber, RNumber, yearFunc(LNumber)))
                    {
                        month = MNumber;
                        day = RNumber;
                        year = yearFunc(LNumber);
                    }
                    else if (legalMonth(MNumber) && legalDay(MNumber, LNumber, yearFunc(RNumber)))
                    {
                        month = MNumber;
                        day = LNumber;
                        year = yearFunc(RNumber);
                    }
                    else
                    {
                        throw new Exception("Invalid date: " + dateLiteral.GetText());
                    }
                }
                else if ((L.DIGIT() == null || M.DIGIT() == null) && R == null)
                {
                    int N;
                    string monthName;
                    if (L.DIGIT() != null)
                    {
                        N = int.Parse(L.GetText());
                        monthName = M.GetText();
                    }
                    else
                    {
                        N = int.Parse(M.GetText());
                        monthName = L.GetText();
                    }
                    int monthNameNumber;
                    if (monthName.Length == 3)
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMM", CultureInfo.InvariantCulture).Month;
                    }
                    else
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMMM", CultureInfo.InvariantCulture).Month;
                    }
                    if (legalDay(monthNameNumber, N, CY))
                    {
                        month = monthNameNumber;
                        day = N;
                        year = CY;
                    }
                    else
                    {
                        month = monthNameNumber;
                        day = 1;
                        year = CY;
                    }
                }
                else
                {
                    int N1;
                    int N2;
                    string monthName;
                    if (L.DIGIT() == null)
                    {
                        monthName = L.GetText();
                        N1 = int.Parse(M.GetText());
                        N2 = int.Parse(R.GetText());
                    }
                    else if (M.DIGIT() == null)
                    {
                        monthName = M.GetText();
                        N1 = int.Parse(L.GetText());
                        N2 = int.Parse(R.GetText());
                    }
                    else
                    {
                        monthName = R.GetText();
                        N1 = int.Parse(L.GetText());
                        N2 = int.Parse(M.GetText());
                    }
                    int monthNameNumber;
                    if (monthName.Length == 3)
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMM", CultureInfo.InvariantCulture).Month;
                    }
                    else
                    {
                        monthNameNumber = DateTime.ParseExact(monthName, "MMMM", CultureInfo.InvariantCulture).Month;
                    }
                    if (legalDay(monthNameNumber, N1, yearFunc(N2)))
                    {
                        month = monthNameNumber;
                        day = N1;
                        year = yearFunc(N2);
                    }
                    else if (legalDay(monthNameNumber, N2, yearFunc(N1)))
                    {
                        month = monthNameNumber;
                        day = N2;
                        year = yearFunc(N1);
                    }
                    else
                    {
                        throw new Exception("Invalid date: " + dateLiteral.GetText());
                    }
                }
            }

            if (dateOrTime.timeValue() == null)
            {
                hours = 0;
                mins = 0;
                seconds = 0;
            }
            else
            {
                var timeValue = dateOrTime.timeValue();
                hours = int.Parse(timeValue.timeValuePart()[0].GetText());
                if (timeValue.timeValuePart().Count == 1)
                {
                    mins = 0;
                }
                else
                {
                    mins = int.Parse(timeValue.timeValuePart()[1].GetText());
                }
                if (timeValue.timeValuePart().Count < 3)
                {
                    seconds = 0;
                }
                else
                {
                    seconds = int.Parse(timeValue.timeValuePart()[2].GetText());
                }
                var amPm = timeValue.AMPM();
                if (amPm != null && (amPm.GetText().ToUpper() == "P" || amPm.GetText().ToUpper() == "PM") && hours >= 0 && hours <= 11)
                {
                    hours += 12;
                }
                else if (amPm != null && (amPm.GetText().ToUpper() == "A" || amPm.GetText().ToUpper() == "AM") && hours == 12)
                {
                    hours = 0;
                }
            }
            var date = new DateTime(year, month, day, hours, mins, seconds);
            return date;
        }

        private object VisitOctLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            string literal = context.OCTLITERAL().GetText();
            literal = literal.Replace("&O", "").Replace("&", "");
            return (decimal)Convert.ToInt32(literal, 8);
        }

        private object VisitHexLiteral(VBAConditionalCompilationParser.LiteralContext context)
        {
            string literal = context.HEXLITERAL().GetText();
            literal = literal.Replace("&H", "").Replace("&", "");
            return (decimal)int.Parse(literal, System.Globalization.NumberStyles.HexNumber);
        }
    }
}
