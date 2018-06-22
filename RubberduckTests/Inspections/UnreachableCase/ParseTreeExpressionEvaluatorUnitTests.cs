using NUnit.Framework;
using Rubberduck.Inspections.Concrete.UnreachableCaseInspection;
using Rubberduck.Parsing.Grammar;
using System;

namespace RubberduckTests.Inspections.UnreachableCase
{
    /*
    ParseTreeValueExpressionEvaluator is a support class of the UnreachableCaseInspection

    Notes: 
    1. The ParseTreeValueExpressionEvaluator uses Longs to work with all Integral types.
    2. The ParseTreeValueExpressionEvaluator uses Double to work with Singles and Doubles.

    1 and 2 mean that, for simplicity, the expressions may be evaluated as a type different than their
    strict, per-spec VBA result type.The goal of the supported inspection is to look for unreachable
    code rather than to precisely replicate VBA's math engine.

    Test Parameter encoding:
    <operand>?<declaredType>_<mathSymbol> _<operand>?<declaredType>, <expression>,<selectExpressionType>
    If there is no "?<declaredType>", then<operand>'s type is derived by the ParseTreeValue instance.
    The<selectExpressionType> is the type that the calculation must yield in order to
    make comparisons in the Select Case statement under inspection.
    */


    [TestFixture]
    public class ParseTreeExpressionEvaluatorUnitTests
    {
        private const string VALUE_TYPE_SEPARATOR = "?";
        private const string OPERAND_SEPARATOR = "_";

        private IUnreachableCaseInspectionFactoryProvider _factoryProvider;
        private IParseTreeValueFactory _valueFactory;
        private IParseTreeExpressionEvaluator _calculator;

        private IUnreachableCaseInspectionFactoryProvider FactoryProvider
        {
            get
            {
                if (_factoryProvider is null)
                {
                    _factoryProvider = new UnreachableCaseInspectionFactoryProvider();
                }
                return _factoryProvider;
            }
        }

        private IParseTreeValueFactory ValueFactory
        {
            get
            {
                if (_valueFactory is null)
                {
                    _valueFactory = FactoryProvider.CreateIParseTreeValueFactory();
                }
                return _valueFactory;
            }
        }

        private IParseTreeExpressionEvaluator Calculator
        {
            get
            {
                if (_calculator is null)
                {
                    _calculator = new ParseTreeExpressionEvaluator(ValueFactory);
                }
                return _calculator;
            }
        }

        [TestCase("x?Byte_-_2?Long", "x - 2", "Long")]
        [TestCase("2_-_x?Byte", "2 - x", "Integer")]
        [TestCase("x?Byte_+_2?Long", "x + 2", "Long")]
        [TestCase("x?Double_/_11.2?Double", "x / 11.2", "Double")]
        [TestCase("x?Double_*_11.2?Double", "x * 11.2", "Double")]
        [TestCase("x?Double_*_y?Double", "x * y", "Double")]
        [TestCase("x?Double_Mod_11.2?Double", "x Mod 11.2", "Double")]
        [TestCase("x?Long_*_y?Double", "x * y", "Double")]
        [TestCase("x?Long_^_11.2?Double", "x ^ 11.2", "Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_VariableMath(string operands, string expected, string selectExpressionTypename)
        {
            GetBinaryOpValues(operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol);
            var result = Calculator.Evaluate(LHS, RHS, opSymbol);
            Assert.AreEqual(result.ValueText, expected);
            Assert.AreEqual(selectExpressionTypename, result.TypeName);
            Assert.IsFalse(result.ParsesToConstantValue);
        }

        [TestCase("-1_>_0", "False", "Boolean")]
        [TestCase("-1.0_>_0.0?Currency", "False", "Boolean")]
        [TestCase("-1_<_0", "True", "Boolean")]
        [TestCase("-1.0_<_0.0?Single", "True", "Boolean")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_RelationalOp(string input, string expected, string selectExpressionTypename)
        {
            GetBinaryOpValues(input, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol);
            var result = Calculator.Evaluate(LHS, RHS, opSymbol);
            Assert.AreEqual(expected, result.ValueText);
        }

        [TestCase("10.51_*_11.2?Currency", "117.712", "Currency")]
        [TestCase("10.51?Currency_*_11.2", "117.712", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_MathOpCurrency(string operands, string expected, string selectExpressionTypename)
        {
            var result = TestBinaryOp(MathSymbols.MULTIPLY, operands, expected, selectExpressionTypename);
            Assert.AreEqual(selectExpressionTypename, result.TypeName);
        }

        [TestCase("10.51?Long_*_11.2", "123.2", "Double")]
        [TestCase("10.51?Integer_*_11.2", "123.2", "Double")]
        [TestCase("10.51?Byte_*_11.2", "123.2", "Double")]
        [TestCase("10.51?Double_*_11.2", "117.712", "Double")]
        [TestCase("10_*_11.2", "112", "Double")]
        [TestCase("11.2_*_10", "112", "Long")]
        [TestCase("10.51_*_11.2", "117.712", "Double")]
        [TestCase("10.51?Single_*_11.2?Single", "117.712", "Single")]
        [TestCase("10.51?Currency_*_11.2?Currency", "117.712", "Single")]
        [TestCase("10_*_11", "110", "Long")]
        [TestCase("True_*_10", "-10", "Long")]
        [TestCase("10_*_True", "-10", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Multiplication(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.MULTIPLY, operands, expected, selectExpressionTypename);
        }

        [TestCase("10_/_2", "5", "Long")]
        [TestCase("2_/_10", "0", "Long")]
        [TestCase("10_/_11.2", "0.89285", "Double")]
        [TestCase("11.2_/_10", "1.12", "Double")]
        [TestCase("10.51_/_11.2", "0.93839286", "Double")]
        [TestCase("10_/_11", "1", "Long")]
        [TestCase(@"""10.51""_/_11.2", "0.93839286", "Double")]
        [TestCase("True_/_10.5", "-0.0952", "Double")]
        [TestCase("10.5_/_True", "-10.5", "Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Division(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.DIVIDE, operands, expected, selectExpressionTypename);
        }

        [TestCase(@"9.5_\_2.4", "5", "Long")]
        [TestCase(@"10_\_4", "2", "Long")]
        [TestCase(@"5.423_\_1", "5", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_IntegerDivision(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.INTEGER_DIVIDE, operands, expected, selectExpressionTypename);
        }

        [TestCase("10.51_+_11.2", "21.71", "Double")]
        [TestCase("10_+_11.2", "21.2", "Double")]
        [TestCase("11.2_+_10", "21.2", "Double")]
        [TestCase("10_+_11", "21", "Long")]
        [TestCase("True_+_10.5", "9.5", "Double")]
        [TestCase("10.5_+_True", "9.5", "Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Addition(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.PLUS, operands, expected, selectExpressionTypename);
        }

        [TestCase("10.51_-_11.2", "-0.69", "Double")]
        [TestCase("10_-_11", "-1", "Long")]
        [TestCase("True_-_10", "-11", "Long")]
        [TestCase("11_-_True", "12", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Subtraction(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.MINUS, operands, expected, selectExpressionTypename);
        }

        [TestCase("10_^_2", "100", "Double")]
        [TestCase("10.5?Currency_^_2.2?Currency", "176.44789", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Powers(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.EXPONENT, operands, expected, selectExpressionTypename);
        }

        [TestCase("10_Mod_3", "1", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Modulo(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(MathSymbols.MODULO, operands, expected, selectExpressionTypename);
        }

        [TestCase("10_=_3", "False")]
        [TestCase("10_=_10", "True")]
        [TestCase("10_<>_3", "True")]
        [TestCase("10_<>_10", "False")]
        [TestCase("10_<_3", "False")]
        [TestCase("10_<=_10", "True")]
        [TestCase("10_>_11", "False")]
        [TestCase("10_>=_10", "True")]
        [TestCase("5_And_0", "False")]
        [TestCase("5_Or_0", "True")]
        [TestCase("5_Xor_6", "False")]
        [TestCase("6.5_>_5.2", "True")]
        [TestCase("True_<_3", "True")]
        [TestCase("False_<_-2", "False")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LogicBinaryConstants(string operands, string expected)
        {
            GetBinaryOpValues(operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase("True_Eqv_True", "True")]
        [TestCase("False_Eqv_True", "False")]
        [TestCase("True_Eqv_False", "False")]
        [TestCase("False_Eqv_False", "True")]
        [TestCase("10_Eqv_8", "-3")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LogicEqvOperator(string operands, string expected)
        {
            GetBinaryOpValues(operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase(true, true, true)]
        [TestCase(false, true, false)]
        [TestCase(true, false, false)]
        [TestCase(false, false, true)]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_EqvOperatorBoolean(bool lhs, bool rhs, bool expected)
        {
            var result = ParseTreeExpressionEvaluator.Eqv(lhs, rhs);
            Assert.AreEqual(expected, result);
        }

        [TestCase(10, 8, -3)]
        [TestCase(3, 0, -4)]
        [TestCase(-1, 0, 0)]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_EqvOperatorInts(int lhs, int rhs, int expected)
        {
            var result = ParseTreeExpressionEvaluator.Eqv(lhs, rhs);
            Assert.AreEqual(expected, result);
        }

        [TestCase(true, true, true)]
        [TestCase(false, true, true)]
        [TestCase(true, false, false)]
        [TestCase(false, false, true)]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_ImpOperatorBoolean(bool lhs, bool rhs, bool expected)
        {
            var result = ParseTreeExpressionEvaluator.Imp(lhs, rhs);
            Assert.AreEqual(expected, result);
        }

        [TestCase(8, 10, -1)]
        [TestCase(3, 0, -4)]
        [TestCase(-1, 0, 0)]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_ImpOperatorInts(int lhs, int rhs, int expected)
        {
            var result = ParseTreeExpressionEvaluator.Imp(lhs, rhs);
            Assert.AreEqual(expected, result);
        }

        [TestCase("8_Imp_10", "-1")]
        [TestCase("True_Imp_True", "True")]
        [TestCase("False_Imp_True", "True")]
        [TestCase("True_Imp_False", "False")]
        [TestCase("False_Imp_False", "True")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LogicImpOperator(string operands, string expected)
        {
            GetBinaryOpValues(operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase("Not_False", "True")]
        [TestCase("Not_True", "False")]
        [TestCase("Not_1", "True")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LogicUnaryConstants(string operands, string expected)
        {
            GetUnaryOpValues(operands, out IParseTreeValue theValue, out string opSymbol);

            var result = Calculator.Evaluate(theValue, opSymbol, Tokens.Boolean);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase("-_45", "-45")]
        [TestCase("-_23.78", "-23.78")]
        [TestCase("-_True", "True?Boolean")]
        [TestCase("-_False", "False?Boolean")]
        [TestCase("-_True", "1?Integer")]
        [TestCase("-_-1", "1?Long")]
        [TestCase("-_0", "False?Boolean")]
        [TestCase("-_1?Double", "-1?Double")]
        [TestCase("-_-1?Double", "1?Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_MinusUnaryOp(string operands, string expected)
        {
            var expectedVal = CreateInspValueFrom(expected);
            GetUnaryOpValues(operands, out IParseTreeValue LHS, out string opSymbol);
            var result = Calculator.Evaluate(LHS, opSymbol, expectedVal.TypeName);

            Assert.AreEqual(expectedVal.ValueText, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }


        [TestCase(@"""A plus sign looks like '+'""_Like_""*+*""", "True")]
        [TestCase(@"""(this is |the d)ay""_Like_""(th*|the*)??""", "True")]
        [TestCase(@"""AB{5} = 25""_Like_""?B{5}*##""", "True")]
        [TestCase(@"""5^2 is 25""_Like_""#^#*##""", "True")]
        [TestCase(@"""$12.50""_Like_""$##[.,]##""", "True")]
        [TestCase(@"""F""_Like_""[a-z]*""", "False")]
        [TestCase(@"""f""_Like_""[a-z]*""", "True")]
        [TestCase(@"""a?bc""_Like_""[a-e][?][a-e]*""", "True")]
        [TestCase(@"""axabc""_Like_""a?abc""", "True")]
        [TestCase(@"""Doggy""_Like_""*Dog*""", "True")]
        [TestCase(@"""Animal""_Like_""[A-Z]*""", "True")]
        [TestCase(@"""123-345-678""_Like_""###[-.]###[-.]###""", "True")]
        [TestCase(@"""BAT123khg""_Like_""B?T*""", "True")]
        [TestCase(@"""CAT123khg""_Like_""B?T*""", "False")]
        [TestCase(@"""aM5b""_Like_""a[L-P]#[!c-e]""", "True")]
        [TestCase(@"""OK?""_Like_""*[?]""", "True")]
        [TestCase(@"""#TryThis""_Like_""[#]*""", "True")]
        [TestCase(@"""#TryThis""_Like_""[#]TryThi?""", "True")]
        [TestCase(@"""**ShineA*SoYouCanSee""_Like_""[*]*Shine*""", "True")]
        [TestCase(@"""FooBard""_Like_""*Bar""", "False")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LikeOperator(string operands, string expected)
        {
            var ops = operands.Split(new string[] { "_" }, StringSplitOptions.None);
            var LHS = ValueFactory.Create(ops[0], Tokens.String);
            var RHS = ValueFactory.Create(ops[2], Tokens.String);
            var result = Calculator.Evaluate(LHS, RHS, ops[1]);

            Assert.AreEqual(expected, result.ValueText, $"{LHS} Like {RHS}");
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase(@"test[[]LBracket", "^test\\[LBracket$")]
        [TestCase(@"[a-e]", "^[a-e]$")]
        [TestCase(@"Bar*", "^Bar[\\D\\d\\s]*")]
        [TestCase(@"[#][1-6]", "^#[1-6]$")]
        [TestCase(@"#[a-e]", "^\\d[a-e]$")]
        [TestCase(@"abc?xy", "^abc.xy$")]
        [TestCase(@"abc[?]xy", "^abc\\?xy$")]
        [TestCase(@"[!A-E]", "^[^A-E]$")]
        [TestCase(@"#[!A-E][#][!5-6][#]*", "^\\d[^A-E]#[^5-6]#[\\D\\d\\s]*")]
        [TestCase(@"abc.xy", "^abc\\.xy$")]
        [TestCase(@"abc[*]xy", "^abc\\*xy$")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LikeRegexConversions(string likePattern, string expectedPattern)
        {
            var result = ParseTreeExpressionEvaluator.ConvertLikePatternToRegexPattern(likePattern);
            Assert.AreEqual(expectedPattern, result);
        }



        [TestCase(@"""Foo""_&_""Bar""", "FooBar")]
        [TestCase(@"1_&_""Bar""", "1Bar")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_AmpersandOperator(string operands, string expected)
        {
            var ops = operands.Split(new string[] { "_" }, StringSplitOptions.None);
            var LHS = ValueFactory.Create(ops[0]);
            var RHS = ValueFactory.Create(ops[2]);
            var result = Calculator.Evaluate(LHS, RHS, ops[1]);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase(@"""A""_=_""A""", "True", true)]
        [TestCase(@"""A""_=_""a""", "False", true)]
        [TestCase(@"""A""_=_""a""", "True", false)]
        [TestCase(@"""A""_<_""a""", "False", true)]
        [TestCase(@"""A""_<_""a""", "False", false)]
        [TestCase(@"""A""_<=_""a""", "True", false)]
        [TestCase(@"""A""_<>_""a""", "True", true)]
        [TestCase(@"""A""_<>_""a""", "False", false)]
        [TestCase(@"""A""_>_""a""", "True", true)]
        [TestCase(@"""A""_>_""a""", "False", false)]
        [TestCase(@"""A""_>=_""a""", "True", true)]
        [TestCase(@"""A""_>=_""a""", "True", false)]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_StringCompares(string operands, string expected, bool optionCompareBinary /*true = caseSensitive*/)
        {
            var ops = operands.Split(new string[] { "_" }, StringSplitOptions.None);
            var LHS = ValueFactory.Create(ops[0]);
            var RHS = ValueFactory.Create(ops[2]);

            var calculator = new ParseTreeExpressionEvaluator(ValueFactory, optionCompareBinary);
            var result = calculator.Evaluate(LHS, RHS, ops[1]);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        //Valid logic operators, but invalid with strings.
        //VBA compiles, but a runtime error would occur - inspection does not flag runtime typemismatch
        [TestCase(@"""A""_And_""a""", "A And a", false)]
        [TestCase(@"""A""_Or_""a""", "A Or a", false)]
        [TestCase(@"""A""_Xor_""a""", "A Xor a", false)]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_StringCompareTypeMismatches(string operands, string expected, bool optionCompareBinary /*true = caseSensitive*/)
        {
            var ops = operands.Split(new string[] { "_" }, StringSplitOptions.None);
            var LHS = ValueFactory.Create(ops[0]);
            var RHS = ValueFactory.Create(ops[2]);

            var calculator = new ParseTreeExpressionEvaluator(ValueFactory, optionCompareBinary);
            var result = calculator.Evaluate(LHS, RHS, ops[1]);

            Assert.AreEqual(expected, result.ValueText);
        }

        private void GetBinaryOpValues(string operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol)
        {
            var operandItems = operands.Split(new string[] { OPERAND_SEPARATOR }, StringSplitOptions.None);

            LHS = CreateInspValueFrom(operandItems[0]);
            opSymbol = operandItems[1];
            RHS = CreateInspValueFrom(operandItems[2]);
        }

        private void GetBinaryOpValues(string operands, string typeName, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol)
        {
            var operandItems = operands.Split(new string[] { OPERAND_SEPARATOR }, StringSplitOptions.None);

            LHS = CreateInspValueFrom(operandItems[0], typeName);
            opSymbol = operandItems[1];
            RHS = CreateInspValueFrom(operandItems[2], typeName);
        }

        private void GetUnaryOpValues(string operands, out IParseTreeValue LHS, out string opSymbol)
        {
            var operandItems = operands.Split(new string[] { OPERAND_SEPARATOR }, StringSplitOptions.None);

            opSymbol = operandItems[0];
            LHS = CreateInspValueFrom(operandItems[1]);
        }

        private IParseTreeValue CreateInspValueFrom(string valAndType, string conformTo = null)
        {
            if (valAndType.Contains(VALUE_TYPE_SEPARATOR))
            {
                var args = valAndType.Split(new string[] { VALUE_TYPE_SEPARATOR }, StringSplitOptions.None);
                var value = args[0];
                string declaredType = args[1].Equals(string.Empty) ? null : args[1];
                if (conformTo is null)
                {
                    if (declaredType is null)
                    {
                        return ValueFactory.Create(value);
                    }
                    return ValueFactory.Create(value, declaredType);
                }
                else
                {
                    if (declaredType is null)
                    {
                        return ValueFactory.Create(value, conformTo);
                    }
                    return ValueFactory.Create(value, declaredType);
                }
            }
            return conformTo is null ? ValueFactory.Create(valAndType)
                : ValueFactory.Create(valAndType, conformTo);
        }

        private IParseTreeValue TestBinaryOp(string opSymbol, string operands, string expected, string selectExpressionTypeName)
        {
            GetBinaryOpValues(operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out _);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            if (selectExpressionTypeName.Equals(Tokens.Double) || selectExpressionTypeName.Equals(Tokens.Single) || selectExpressionTypeName.Equals(Tokens.Currency))
            {
                var compareLength = expected.Length > 5 ? 5 : expected.Length;
                Assert.IsTrue(Math.Abs(double.Parse(result.ValueText.Substring(0, compareLength)) - double.Parse(expected.Substring(0, compareLength))) <= double.Epsilon, $"Actual={result.ValueText} Expected={expected}");
            }
            else if (selectExpressionTypeName.Equals(Tokens.String))
            {
                var compareLength = expected.Length > 5 ? 5 : expected.Length;
                Assert.AreEqual(expected.Substring(0, compareLength), result.ValueText.Substring(0, compareLength));
            }
            else
            {
                Assert.AreEqual(expected, result.ValueText);
            }
            Assert.IsTrue(result.ParsesToConstantValue, "Expected 'IsConstantValue' property to be true");
            return result;
        }
    }
}
