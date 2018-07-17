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

        [TestCase("Boolean", "Byte", "Integer")]
        [TestCase("Boolean", "Boolean", "Integer")]
        [TestCase("Boolean", "Integer", "Integer")]
        [TestCase("Integer", "Byte", "Integer")]
        [TestCase("Integer", "Boolean", "Integer")]
        [TestCase("Integer", "Integer", "Integer")]
        [TestCase("Byte", "Byte", "Byte")]
        [TestCase("Byte", "Boolean", "Integer")]
        [TestCase("Byte", "Integer", "Integer")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticIntegerType(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Long", "Byte", "Long")]
        [TestCase("Long", "Boolean", "Long")]
        [TestCase("Long", "Integer", "Long")]
        [TestCase("Long", "Long", "Long")]
        [TestCase("Byte", "Long", "Long")]
        [TestCase("Boolean", "Long", "Long")]
        [TestCase("Integer", "Long", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticLong(string lhsType, string rhsType, string expectedResultType)
        {
            var opType = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(opType.OperatorDeclaredType, expectedResultType);
        }

        [TestCase("LongLong", "Byte", "LongLong")]
        [TestCase("LongLong", "Boolean", "LongLong")]
        [TestCase("LongLong", "Integer", "LongLong")]
        [TestCase("LongLong", "Long", "LongLong")]
        [TestCase("LongLong", "LongLong", "LongLong")]
        [TestCase("Byte", "LongLong", "LongLong")]
        [TestCase("Boolean", "LongLong", "LongLong")]
        [TestCase("Integer", "LongLong", "LongLong")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticLongLong(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Single", "Byte", "Single")]
        [TestCase("Single", "Boolean", "Single")]
        [TestCase("Single", "Integer", "Single")]
        [TestCase("Single", "Single", "Single")]
        [TestCase("Byte", "Single", "Single")]
        [TestCase("Boolean", "Single", "Single")]
        [TestCase("Integer", "Single", "Single")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticSingle(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }


        [TestCase("Double", "Double", "Double")]
        [TestCase("Double", "String", "Double")]
        [TestCase("Double", "Byte", "Double")]
        [TestCase("Double", "Boolean", "Double")]
        [TestCase("Double", "Integer", "Double")]
        [TestCase("String", "Double", "Double")]
        [TestCase("String", "String", "Double")]
        [TestCase("String", "Byte", "Double")]
        [TestCase("String", "Boolean", "Double")]
        [TestCase("String", "Integer", "Double")]
        [TestCase("Long", "Double", "Double")]
        [TestCase("Integer", "Double", "Double")]
        [TestCase("Byte", "Double", "Double")]
        [TestCase("Boolean", "Double", "Double")]
        [TestCase("Single", "Long", "Double")]
        [TestCase("Single", "LongLong", "Double")]
        [TestCase("Long", "Single", "Double")]
        [TestCase("LongLong", "Single", "Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticDouble(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Currency", "Byte", "Currency")]
        [TestCase("Currency", "Boolean", "Currency")]
        [TestCase("Currency", "Integer", "Currency")]
        [TestCase("Currency", "Long", "Currency")]
        [TestCase("Currency", "LongLong", "Currency")]
        [TestCase("Currency", "Double", "Currency")]
        [TestCase("Currency", "Single", "Currency")]
        [TestCase("Currency", "Currency", "Currency")]
        [TestCase("Byte", "Currency", "Currency")]
        [TestCase("Boolean", "Currency", "Currency")]
        [TestCase("Integer", "Currency", "Currency")]
        [TestCase("Long", "Currency", "Currency")]
        [TestCase("LongLong", "Currency", "Currency")]
        [TestCase("Double", "Currency", "Currency")]
        [TestCase("Single", "Currency", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticCurrency(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Date", "Byte", "Date")]
        [TestCase("Date", "Boolean", "Date")]
        [TestCase("Date", "Integer", "Date")]
        [TestCase("Date", "Long", "Date")]
        [TestCase("Date", "LongLong", "Date")]
        [TestCase("Date", "Double", "Date")]
        [TestCase("Date", "Single", "Date")]
        [TestCase("Date", "Currency", "Date")]
        [TestCase("Date", "String", "Date")]
        [TestCase("Date", "Date", "Date")]
        [TestCase("Byte", "Date", "Date")]
        [TestCase("Boolean", "Date", "Date")]
        [TestCase("Integer", "Date", "Date")]
        [TestCase("Long", "Date", "Date")]
        [TestCase("LongLong", "Date", "Date")]
        [TestCase("Double", "Date", "Date")]
        [TestCase("Single", "Date", "Date")]
        [TestCase("Currency", "Date", "Date")]
        [TestCase("String", "Date", "Date")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticDate(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Arithmetic);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Currency", "Double", "Double")]
        [TestCase("Currency", "Single", "Double")]
        [TestCase("Currency", "String", "Double")]
        [TestCase("Double", "Currency", "Double")]
        [TestCase("Single", "Currency", "Double")]
        [TestCase("String", "Currency", "Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticMult(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), ArithmeticOperators.MULTIPLY);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Byte", "Byte", "Double")]
        [TestCase("Byte", "Boolean", "Double")]
        [TestCase("Byte", "Integer", "Double")]
        [TestCase("Byte", "Long", "Double")]
        [TestCase("Byte", "LongLong", "Double")]
        [TestCase("Boolean", "Byte", "Double")]
        [TestCase("Boolean", "Boolean", "Double")]
        [TestCase("Boolean", "Integer", "Double")]
        [TestCase("Boolean", "Long", "Double")]
        [TestCase("Boolean", "LongLong", "Double")]
        [TestCase("Integer", "Byte", "Double")]
        [TestCase("Integer", "Boolean", "Double")]
        [TestCase("Integer", "Integer", "Double")]
        [TestCase("Integer", "Long", "Double")]
        [TestCase("Integer", "LongLong", "Double")]
        [TestCase("Long", "Byte", "Double")]
        [TestCase("Long", "Boolean", "Double")]
        [TestCase("Long", "Integer", "Double")]
        [TestCase("Long", "Long", "Double")]
        [TestCase("Long", "LongLong", "Double")]
        [TestCase("LongLong", "Byte", "Double")]
        [TestCase("LongLong", "Boolean", "Double")]
        [TestCase("LongLong", "Integer", "Double")]
        [TestCase("LongLong", "Long", "Double")]
        [TestCase("LongLong", "LongLong", "Double")]
        [TestCase("Byte", "Byte", "Double")]
        [TestCase("Double", "Boolean", "Double")]
        [TestCase("Double", "Integer", "Double")]
        [TestCase("Double", "Long", "Double")]
        [TestCase("Double", "LongLong", "Double")]
        [TestCase("Double", "Single", "Double")]
        [TestCase("Double", "Double", "Double")]
        [TestCase("Double", "Currency", "Double")]
        [TestCase("Double", "String", "Double")]
        [TestCase("Double", "Date", "Double")]
        [TestCase("String", "Boolean", "Double")]
        [TestCase("String", "Integer", "Double")]
        [TestCase("String", "Long", "Double")]
        [TestCase("String", "LongLong", "Double")]
        [TestCase("String", "Single", "Double")]
        [TestCase("String", "Double", "Double")]
        [TestCase("String", "Currency", "Double")]
        [TestCase("String", "String", "Double")]
        [TestCase("String", "Date", "Double")]
        [TestCase("Currency", "Boolean", "Double")]
        [TestCase("Currency", "Integer", "Double")]
        [TestCase("Currency", "Long", "Double")]
        [TestCase("Currency", "LongLong", "Double")]
        [TestCase("Currency", "Single", "Double")]
        [TestCase("Currency", "Double", "Double")]
        [TestCase("Currency", "Currency", "Double")]
        [TestCase("Currency", "String", "Double")]
        [TestCase("Currency", "Date", "Double")]
        [TestCase("Date", "Boolean", "Double")]
        [TestCase("Date", "Integer", "Double")]
        [TestCase("Date", "Long", "Double")]
        [TestCase("Date", "LongLong", "Double")]
        [TestCase("Date", "Single", "Double")]
        [TestCase("Date", "Double", "Double")]
        [TestCase("Date", "Currency", "Double")]
        [TestCase("Date", "String", "Double")]
        [TestCase("Date", "Date", "Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticDiv(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), ArithmeticOperators.DIVIDE);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }

        [TestCase("Currency", "Boolean", "Long")]
        [TestCase("Currency", "Integer", "Long")]
        [TestCase("Currency", "Long", "Long")]
        [TestCase("Currency", "LongLong", "Long")]
        [TestCase("Currency", "Single", "Long")]
        [TestCase("Currency", "Double", "Long")]
        [TestCase("Currency", "Currency", "Long")]
        [TestCase("Currency", "String", "Long")]
        [TestCase("Currency", "Date", "Long")]
        [TestCase("Double", "Boolean", "Long")]
        [TestCase("Double", "Integer", "Long")]
        [TestCase("Double", "Long", "Long")]
        [TestCase("Double", "LongLong", "Long")]
        [TestCase("Double", "Single", "Long")]
        [TestCase("Double", "Double", "Long")]
        [TestCase("Double", "Currency", "Long")]
        [TestCase("Double", "String", "Long")]
        [TestCase("Double", "Date", "Long")]
        [TestCase("Single", "Boolean", "Long")]
        [TestCase("Single", "Integer", "Long")]
        [TestCase("Single", "Long", "Long")]
        [TestCase("Single", "LongLong", "Long")]
        [TestCase("Single", "Single", "Long")]
        [TestCase("Single", "Double", "Long")]
        [TestCase("Single", "Currency", "Long")]
        [TestCase("Single", "String", "Long")]
        [TestCase("Single", "Date", "Long")]
        [TestCase("String", "Boolean", "Long")]
        [TestCase("String", "Integer", "Long")]
        [TestCase("String", "Long", "Long")]
        [TestCase("String", "LongLong", "Long")]
        [TestCase("String", "Single", "Long")]
        [TestCase("String", "Double", "Long")]
        [TestCase("String", "Currency", "Long")]
        [TestCase("String", "String", "Long")]
        [TestCase("String", "Date", "Long")]
        [TestCase("Date", "Boolean", "Long")]
        [TestCase("Date", "Integer", "Long")]
        [TestCase("Date", "Long", "Long")]
        [TestCase("Date", "LongLong", "Long")]
        [TestCase("Date", "Single", "Long")]
        [TestCase("Date", "Double", "Long")]
        [TestCase("Date", "Currency", "Long")]
        [TestCase("Date", "String", "Long")]
        [TestCase("Date", "Date", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeArithmeticMod(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), ArithmeticOperators.MODULO);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType );
        }

        [TestCase("Byte", "Byte", "Byte")]
        [TestCase("Boolean", "Boolean", "Boolean")]
        [TestCase("Integer", "Integer", "Integer")]
        [TestCase("Currency", "Byte", "Long")]
        [TestCase("Currency", "Boolean", "Long")]
        [TestCase("Currency", "Integer", "Long")]
        [TestCase("Currency", "Long", "Long")]
        [TestCase("Currency", "Double", "Long")]
        [TestCase("Currency", "Single", "Long")]
        [TestCase("Currency", "Currency", "Long")]
        [TestCase("Double", "Byte", "Long")]
        [TestCase("Double", "Boolean", "Long")]
        [TestCase("Double", "Integer", "Long")]
        [TestCase("Double", "Long", "Long")]
        [TestCase("Double", "Double", "Long")]
        [TestCase("Double", "Single", "Long")]
        [TestCase("Double", "Currency", "Long")]
        [TestCase("Single", "Long", "Long")]
        [TestCase("Long", "Byte", "Long")]
        [TestCase("Long", "Boolean", "Long")]
        [TestCase("Long", "Integer", "Long")]
        [TestCase("Long", "Long", "Long")]
        [TestCase("Long", "Double", "Long")]
        [TestCase("Long", "Single", "Long")]
        [TestCase("Long", "Currency", "Long")]
        [TestCase("String", "Byte", "Long")]
        [TestCase("String", "Boolean", "Long")]
        [TestCase("String", "Integer", "Long")]
        [TestCase("String", "Long", "Long")]
        [TestCase("String", "Double", "Long")]
        [TestCase("String", "Single", "Long")]
        [TestCase("String", "Currency", "Long")]
        [TestCase("Date", "Byte", "Long")]
        [TestCase("Date", "Boolean", "Long")]
        [TestCase("Date", "Integer", "Long")]
        [TestCase("Date", "Long", "Long")]
        [TestCase("Date", "Double", "Long")]
        [TestCase("Date", "Single", "Long")]
        [TestCase("Date", "Currency", "Long")]
        [TestCase("LongLong", "Byte", "LongLong")]
        [TestCase("LongLong", "Boolean", "LongLong")]
        [TestCase("LongLong", "Integer", "LongLong")]
        [TestCase("LongLong", "Long", "LongLong")]
        [TestCase("LongLong", "Double", "LongLong")]
        [TestCase("LongLong", "Single", "LongLong")]
        [TestCase("LongLong", "Currency", "LongLong")]
        [TestCase("LongLong", "String", "LongLong")]
        [TestCase("LongLong", "Date", "LongLong")]
        [TestCase("Byte", "Variant", "Variant")]
        [TestCase("Boolean", "Variant", "Variant")]
        [TestCase("Integer", "Variant", "Variant")]
        [TestCase("Long", "Variant", "Variant")]
        [TestCase("Double", "Variant", "Variant")]
        [TestCase("Single", "Variant", "Variant")]
        [TestCase("Currency", "Variant", "Variant")]
        [TestCase("String", "Variant", "Variant")]
        [TestCase("Date", "Variant", "Variant")]
        [TestCase("Variant", "Variant", "Variant")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_OperatorResultTypeLogical(string lhsType, string rhsType, string expectedResultType)
        {
            var opProvider = new OperatorDeclaredTypeProvider((lhsType, rhsType), OperatorDeclaredTypeProviderTypes.Logical);
            Assert.AreEqual(expectedResultType, opProvider.OperatorDeclaredType);
        }
        [TestCase("x?Byte_-_2?Long", "x - 2", "Long")]
        [TestCase("2_-_x?Byte", "2 - x", "Integer")]
        [TestCase("x?Byte_+_2?Long", "x + 2", "Long")]
        [TestCase("x?Double_/_11.2?Double", "x / 11.2", "Double")]
        [TestCase("x?Double_*_11.2?Double", "x * 11.2", "Double")]
        [TestCase("x?Double_*_y?Double", "x * y", "Double")]
        [TestCase("x?Double_Mod_11.2?Double", "x Mod 11.2", "Long")]
        [TestCase("x?Double_\\_11.2?Double", "x \\ 11.2", "Long")]
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

        [TestCase("10.51_*_11.2?Currency", "117.712", "Double")]
        [TestCase("10.51?Currency_*_11.2", "117.712", "Double")]
        [TestCase("10.51?Currency_*_11.2?Currency", "117.712", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_MathOpCurrency(string operands, string expected, string selectExpressionTypename)
        {
            var result = TestBinaryOp(ArithmeticOperators.MULTIPLY, operands, expected, selectExpressionTypename);
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
            TestBinaryOp(ArithmeticOperators.MULTIPLY, operands, expected, selectExpressionTypename);
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
            TestBinaryOp(ArithmeticOperators.DIVIDE, operands, expected, selectExpressionTypename);
        }

        [TestCase(@"9.5_\_2.4", "5", "Long")]
        [TestCase(@"10_\_4", "2", "Long")]
        [TestCase(@"5.423_\_1", "5", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_IntegerDivision(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(ArithmeticOperators.INTEGER_DIVIDE, operands, expected, selectExpressionTypename);
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
            TestBinaryOp(ArithmeticOperators.PLUS, operands, expected, selectExpressionTypename);
        }

        [TestCase("10.51_-_11.2", "-0.69", "Double")]
        [TestCase("10_-_11", "-1", "Long")]
        [TestCase("True_-_10", "-11", "Long")]
        [TestCase("11_-_True", "12", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Subtraction(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(ArithmeticOperators.MINUS, operands, expected, selectExpressionTypename);
        }

        [TestCase("10_^_2", "100", "Double")]
        [TestCase("10.5?Currency_^_2.2?Currency", "176.44789", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Powers(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(ArithmeticOperators.EXPONENT, operands, expected, selectExpressionTypename);
        }

        [TestCase("10_Mod_3", "1", "Currency")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_Modulo(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(ArithmeticOperators.MODULO, operands, expected, selectExpressionTypename);
        }

        [TestCase("5_And_1", "1")]
        [TestCase("5_And_0", "0")]
        [TestCase("5_Or_0", "5")]
        [TestCase("5_Xor_6", "3")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LogicBinaryConstants(string operands, string expected)
        {
            GetBinaryOpValues(operands, out IParseTreeValue LHS, out IParseTreeValue RHS, out string opSymbol);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase("10_=_3", "False")]
        [TestCase("10_=_10", "True")]
        [TestCase("10_<>_3", "True")]
        [TestCase("10_<>_10", "False")]
        [TestCase("10_<_3", "False")]
        [TestCase("10_<=_10", "True")]
        [TestCase("10_=<_10", "True")]
        [TestCase("10_>_11", "False")]
        [TestCase("10_>=_10", "True")]
        [TestCase("10_=>_10", "True")]
        [TestCase("6.5_>_5.2", "True")]
        [TestCase("True_<_3", "True")]
        [TestCase("False_<_-2", "False")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_RelationalOpConstants(string operands, string expected)
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
        [TestCase("False_Eqv_5", "-6")]
        [TestCase("True_Eqv_5", "5")]
        [TestCase("5_Eqv_False", "-6")]
        [TestCase("5_Eqv_True", "5")]
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

        [TestCase("8_Imp_10", "-1")]
        [TestCase("3_Imp_0", "-4")]
        [TestCase("-1_Imp_0", "0")]
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
        [TestCase("Not_1", "-2")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_LogicUnaryConstants(string operands, string expected)
        {
            GetUnaryOpValues(operands, out IParseTreeValue theValue, out string opSymbol);

            var result = Calculator.Evaluate(theValue, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase("-_45", "-45")]
        [TestCase("-_23.78", "-23.78")]
        [TestCase("-_True", "1?Integer")]
        [TestCase("-_False", "0?Integer")]
        [TestCase("-_True", "1?Integer")]
        [TestCase("-_-1", "1?Integer")]
        [TestCase("-_0", "0?Integer")]
        [TestCase("-_1?Double", "-1?Double")]
        [TestCase("-_-1?Double", "1?Double")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_MinusUnaryOp(string operands, string expected)
        {
            var expectedVal = CreateInspValueFrom(expected);
            GetUnaryOpValues(operands, out IParseTreeValue LHS, out string opSymbol);
            var result = Calculator.Evaluate(LHS, opSymbol);

            Assert.AreEqual(expectedVal.ValueText, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
            Assert.AreEqual(expectedVal.TypeName, result.TypeName);
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

        [TestCase(@"""2""_+_""2""", "22", "String")]
        [TestCase(@"""2""_+_2", "4", "Long")]
        [TestCase(@"""256""_+_""2""", "2562", "Long")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_AdditionString(string operands, string expected, string selectExpressionTypename)
        {
            TestBinaryOp(ArithmeticOperators.PLUS, operands, expected, selectExpressionTypename);
        }

        [TestCase("<", ">")]
        [TestCase(">", "<")]
        [TestCase(">=", "<=")]
        [TestCase("<=", ">=")]
        [TestCase("=", "=")]
        [TestCase("<>", "<>")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_PredicateMatchesSelectExpression(string initialSign, string invertedSign)
        {
            var selectExpression = "x";
            var variableExpression = "z";
            var lhs = ValueFactory.Create(variableExpression);
            var rhs = ValueFactory.Create(selectExpression);
            var symbol = new Tuple<string, string>(initialSign, invertedSign);
            var predicate = new BinaryExpression(lhs, rhs, symbol.Item1);
            var expected = $"{selectExpression} {symbol.Item2} {variableExpression}";
            Assert.AreEqual(expected, predicate.ToString());
        }

        [TestCase("45", "<", "x > 45")]
        [TestCase("45", "And", "x And 45")]
        [TestCase("z", "<", "x > z")]
        [TestCase("z", "Or", "x Or z")]
        [TestCase("z", "Xor", "x Xor z")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_PredicateMovesVariablesLeft(string input, string symbol, string expected)
        {
            var variableExpression = "x";
            var lhs = ValueFactory.Create(input);
            var rhs = ValueFactory.Create(variableExpression);
            var predicate = new BinaryExpression(lhs, rhs, symbol);
            Assert.AreEqual(expected, predicate.ToString());
        }

        [TestCase("Eqv")]
        [TestCase("Imp")]
        [TestCase("Like")]
        [Category("Inspections")]
        public void ParseTreeValueExpressionEvaluator_PredicateNoAlgebra(string symbol)
        {
            var input = "45";
            var selectExpression = "x";
            var expected = $"{input} {symbol} {selectExpression}";
            var lhs = ValueFactory.Create(input);
            var rhs = ValueFactory.Create(selectExpression);
            var predicate = new BinaryExpression(lhs, rhs, symbol);
            Assert.AreEqual(expected, predicate.ToString());
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
                if (result.TryConvertValue(out long selectExpressionResult))
                {
                    Assert.AreEqual(expected, selectExpressionResult.ToString());
                }
                else
                {
                    Assert.Fail("Unable to convert result to Select Expression Type");
                }
            }
            Assert.IsTrue(result.ParsesToConstantValue, "Expected 'IsConstantValue' property to be true");
            return result;
        }
    }
}
