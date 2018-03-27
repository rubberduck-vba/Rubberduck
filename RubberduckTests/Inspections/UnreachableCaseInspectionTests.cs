using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete.UnreachableCaseInspection;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnreachableCaseInspectionTests
    {
        private const string VALUE_TYPE_SEPARATOR = "?";
        private const string OPERAND_SEPARATOR = "_";

        private IUnreachableCaseInspectionFactoryFactory _factoriesFactory;
        private IUCIValueFactory _valueFactory;
        private IUCIValueExpressionEvaluator _calculator;
        private IUCIParseTreeValueVisitorFactory _visitorFactory;
        private IUCIRangeClauseFilterFactory _rangeClauseFilterFactory;
        private IUnreachableCaseInspectionRangeFactory _rangeFactory;
        private IUnreachableCaseInspectionSelectStmtFactory _selectStmtFactory;
        private Dictionary<ParserRuleContext, IUCIValue> _inspectionResults;

        private void Test_OnValueResultCreated(object sender, ValueResultEventArgs e)
        {
            ParseValueResults.Add(e.Context, e.Value);
        }

        private Dictionary<ParserRuleContext, IUCIValue> ParseValueResults
        {
            get
            {
                if(_inspectionResults is null)
                {
                    _inspectionResults = new Dictionary<ParserRuleContext, IUCIValue>();
                }
                return _inspectionResults;
            }
        }

        private IUnreachableCaseInspectionFactoryFactory FactoriesFactoryTest
        {
            get
            {
                if (_factoriesFactory is null)
                {
                    _factoriesFactory = new UnreachableCaseInspectionFactoryFactory();
                }
                return _factoriesFactory;
             }
        }

        private IUCIValueFactory ValueFactory
        {
            get
            {
                if(_valueFactory is null)
                {
                    _valueFactory = FactoriesFactoryTest.CreateIUCIValueFactory();
                }
                return _valueFactory;
            }
        }

        private IUCIValueExpressionEvaluator Calculator
        {
            get
            {
                if (_calculator is null)
                {
                    _calculator = new UCIValueExpressionEvaluator(ValueFactory);
                }
                return _calculator;
            }
        }

        private IUCIParseTreeValueVisitorFactory ValueVisitorFactory
        {
            get
            {
                if (_visitorFactory is null)
                {
                    _visitorFactory = FactoriesFactoryTest.CreateIUCIParseTreeValueVisitorFactory();
                }
                return _visitorFactory;
            }
        }

        private IUCIRangeClauseFilterFactory RangeClauseFilterFactory
        {
            get
            {
                if (_rangeClauseFilterFactory is null)
                {
                    _rangeClauseFilterFactory = FactoriesFactoryTest.CreateIUCIRangeClauseFilterFactory();
                }
                return _rangeClauseFilterFactory;
            }
        }

        private IUnreachableCaseInspectionRangeFactory InspectionRangeFactory
        {
            get
            {
                if (_rangeFactory is null)
                {
                    _rangeFactory = FactoriesFactoryTest.CreateUnreachableCaseInspectionRangeFactory();
                }
                return _rangeFactory;
            }
        }

        private IUnreachableCaseInspectionSelectStmtFactory InspectionSelectStmtFactory
        {
            get
            {
                if (_selectStmtFactory is null)
                {
                    _selectStmtFactory = FactoriesFactoryTest.CreateUnreachableCaseInspectionSelectStmtFactory();
                }
                return _selectStmtFactory;
            }
        }

        [TestCase("2", "2")]
        [TestCase("2.54", "2.54")]
        [TestCase("2.54?Long", "3")]
        [TestCase("2.54?Double", "2.54")]
        [TestCase("2.54?Boolean", "True")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciConformedTypes(string operands, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciNullInputValue()
        {
            IUCIValue test = null;
            try
            {
                test = ValueFactory.Create(null);
                Assert.Fail("Null input to UnreachableCaseInspectionValue did not generate an Argument exception");
            }
            catch (ArgumentException)
            {

            }
            catch
            {
                Assert.Fail("Null input to UnreachableCaseInspectionValue did not generate an exception");
            }
        }

        [TestCase("x", "","x")]
        [TestCase("x?Variant", "Variant", "x")]
        [TestCase("x?String", "String", "x")]
        [TestCase("x?Double","Double", "x")]
        [TestCase("x456", "", "x456")]
        [TestCase(@"""x456""", "String", "x456")]
        [TestCase("x456?String", "String", "x456")]
        [TestCase("45E2", "Double", "4500")]
        [TestCase(@"""10.51""", "String","10.51")]
        [TestCase(@"""What@""", "String","What@")]
        [TestCase(@"""What!""", "String","What!")]
        [TestCase(@"""What#""", "String","What#")]
        [TestCase("What%", "Integer","What")]
        [TestCase("What&", "Long","What")]
        [TestCase("What@", "Decimal","What")]
        [TestCase("What!", "Single", "What")]
        [TestCase("What#", "Double", "What")]
        [TestCase("What$", "String", "What")]
        [TestCase("345%", "Integer","345")]
        [TestCase("45#", "Double", "45")]
        [TestCase("45@", "Decimal", "45")]
        [TestCase("45!", "Single", "45")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciVariableTypes(string operands, string expectedTypeName, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands);
            Assert.AreEqual(expectedTypeName, value.TypeName);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [TestCase("45.5?Double", "Double", "45.5")]
        [TestCase("45.5?Currency", "Currency", "45.5")]
        [TestCase(@"""45E2""?Long", "Long", "4500")]
        [TestCase(@"""95E-2""?Double", "Double", "0.95")]
        [TestCase(@"""95E-2""?Byte", "Byte", "1")]
        [TestCase("True?Double", "Double", "-1")]
        [TestCase("True?Long", "Long", "-1")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciConformToType(string operands, string conformToType, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands, conformToType);
            
            Assert.AreEqual(conformToType, value.TypeName);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [TestCase("x?Byte_-_2?Long", "x - 2", "Long")]
        [TestCase("2_-_x?Byte?Long", "2 - x", "Long")]
        [TestCase("x?Byte_+_2?Long", "x + 2", "Long")]
        [TestCase("x?Double_/_11.2?Double", "x / 11.2", "Double")]
        [TestCase("x?Double_*_11.2?Double", "x * 11.2", "Double")]
        [TestCase("x?Double_*_y?Double", "x * y", "Double")]
        [TestCase("x?Double_Mod_11.2?Double", "x Mod 11.2", "Double")]
        [TestCase("x?Long_*_y?Double", "x * y", "Double")]
        [TestCase("x?Long_^_11.2?Double", "x ^ 11.2", "Double")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciVariableMath(string operands, string expected, string typeName)
        {
            GetBinaryOpValues(operands, out IUCIValue LHS, out IUCIValue RHS, out string opSymbol);
            var result = Calculator.Evaluate(LHS, RHS, opSymbol);
            Assert.AreEqual(result.ValueText, expected);
            Assert.AreEqual(typeName, result.TypeName);
            Assert.IsFalse(result.ParsesToConstantValue, "ConstantValue field expected to be false");
        }

        [TestCase("-1_>_0", "False", "Boolean")]
        [TestCase("-1.0_>_0.0?Currency", "False", "Boolean")]
        [TestCase("-1_<_0", "True", "Boolean")]
        [TestCase("-1.0_<_0.0?Single", "True", "Boolean")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciRelationalOp(string input, string expected, string typeName)
        {
            GetBinaryOpValues(input, out IUCIValue LHS, out IUCIValue RHS, out string opSymbol);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);
            Assert.AreEqual(expected, result.ValueText);
        }

        [TestCase("False", "False")]
        [TestCase("True", "True")]
        [TestCase("-1", "True")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciConvertToBoolText(string input, string expected)
        {
            var result = ValueFactory.Create(input, Tokens.Boolean);
            Assert.IsNotNull(result, $"Type conversion to {Tokens.Boolean} return null interface");
            Assert.AreEqual(expected, result.ValueText);
        }

        [TestCase("Yahoo", "Long")]
        [TestCase("Yahoo", "Double")]
        [TestCase("Yahoo", "Boolean")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciConvertToType(string input, string convertToTypeName)
        {
            var result =ValueFactory.Create(input, convertToTypeName);
            Assert.IsNotNull(result, $"Type conversion to {convertToTypeName} return null interface");
            Assert.AreEqual("Yahoo", result.ValueText);
        }

        [TestCase("NaN", "String")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciConvertToNanType(string input, string convertToTypeName)
        {
            var result = ValueFactory.Create(input, convertToTypeName);
            Assert.IsNotNull(result, $"Type conversion to {convertToTypeName} return null interface");
            Assert.AreEqual("NaN", result.ValueText);
        }

        [TestCase("10.51_*_11.2?Currency", "117.712", "Currency")]
        [TestCase("10.51?Currency_*_11.2", "117.712", "Currency")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciHandlesCurrency(string operands, string expected, string typeName)
        {
            var result = TestBinaryOp(MathTokens.MULT, operands, expected, typeName);
            Assert.AreEqual(typeName, result.TypeName);
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
        [TestCase("10_*_11", "110", "long")]
        [TestCase("True_*_10", "-10", "Long")]
        [TestCase("10_*_True", "-10", "Long")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciMultiplication(string operands, string expected, string typeName)
        {
            TestBinaryOp(MathTokens.MULT, operands, expected, typeName);
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
        public void UnreachableCaseInspUnit_uciDivision(string operands, string expected, string typeName)
        {
            TestBinaryOp(MathTokens.DIV, operands, expected, typeName);
        }

        [TestCase("10.51_+_11.2", "21.71", "Double")]
        [TestCase("10_+_11.2", "21.2", "Double")]
        [TestCase("11.2_+_10", "21.2", "Double")]
        [TestCase("10_+_11", "21", "Long")]
        [TestCase("True_+_10.5", "9.5", "Double")]
        [TestCase("10.5_+_True", "9.5", "Double")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciAddition(string operands, string expected, string typeName)
        {
            TestBinaryOp(MathTokens.ADD, operands, expected, typeName);
        }

        [TestCase("10.51_-_11.2", "-0.69", "Double")]
        [TestCase("10_-_11", "-1", "Long")]
        [TestCase("True_-_10", "-11", "Long")]
        [TestCase("11_-_True", "12", "Long")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciSubtraction(string operands, string expected, string typeName)
        {
            TestBinaryOp(MathTokens.SUBTRACT, operands, expected, typeName);
        }

        [TestCase("10_^_2", "100", "Double")]
        [TestCase("10.5?Currency_^_2.2?Currency", "176.44789", "Currency")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciPowers(string operands, string expected, string typeName)
        {
            TestBinaryOp(MathTokens.POW, operands, expected, typeName);
        }

        [TestCase("10_Mod_3", "1", "Currency")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciModulo(string operands, string expected, string typeName)
        {
            TestBinaryOp(MathTokens.MOD, operands, expected, typeName);
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
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciLogicBinaryConstants(string operands, string expected)
        {
            GetBinaryOpValues(operands, out IUCIValue LHS, out IUCIValue RHS, out string opSymbol);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        [TestCase("Not_False", "True")]
        [TestCase("Not_True", "False")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciLogicUnaryConstants(string operands, string expected)
        {
            GetUnaryOpValues(operands, out IUCIValue theValue, out string opSymbol);

            var result = Calculator.Evaluate(theValue, opSymbol);

            Assert.AreEqual(expected, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue, "Expected IsConstantValue field to be 'True'");
        }

        [TestCase("45", "-45")]
        [TestCase("23.78", "-23.78")]
        [TestCase("True", "True?Boolean")]
        [TestCase("False", "False?Boolean")]
        [TestCase("-1?Double", "1?Double")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciMinusUnaryOp(string operands, string expected)
        {
            var theValue = CreateInspValueFrom(operands);
            var expectedVal = CreateInspValueFrom(expected);
            var opSymbol = MathTokens.SUBTRACT;

            var result = Calculator.Evaluate(theValue, opSymbol);

            Assert.AreEqual(expectedVal.ValueText, result.ValueText);
            Assert.IsTrue(result.ParsesToConstantValue);
        }

        private IUCIValue TestBinaryOp(string opSymbol, string operands, string expected, string typeName)
        {
            GetBinaryOpValues(operands, out IUCIValue LHS, out IUCIValue RHS, out _);

            var result = Calculator.Evaluate(LHS, RHS, opSymbol);

            if (typeName.Equals(Tokens.Double) || typeName.Equals(Tokens.Single) || typeName.Equals(Tokens.Currency))
            {
                Assert.IsTrue(Math.Abs(double.Parse(result.ValueText) - double.Parse(expected)) < .001, $"Actual={result.ValueText} Expected={expected}");
            }
            else if (typeName.Equals(Tokens.String))
            {
                var toComp = expected.Length > 5 ? 5 : expected.Length;
                Assert.AreEqual(expected.Substring(0, toComp), result.ValueText.Substring(0, toComp));
            }
            else
            {
                Assert.AreEqual(expected, result.ValueText);
            }
            Assert.IsTrue(result.ParsesToConstantValue, "Expected 'IsConstantValue' property to be true");
            return result;
        }

        private void GetBinaryOpValues(string operands, out IUCIValue LHS, out IUCIValue RHS, out string opSymbol)
        {
            var operandItems = operands.Split(new string[] { OPERAND_SEPARATOR }, StringSplitOptions.None);

            LHS = CreateInspValueFrom(operandItems[0]);
            opSymbol = operandItems[1];
            RHS = CreateInspValueFrom(operandItems[2]);
        }

        private void GetUnaryOpValues(string operands, out IUCIValue LHS, out string opSymbol)
        {
            var operandItems = operands.Split(new string[] { OPERAND_SEPARATOR }, StringSplitOptions.None);

            opSymbol = operandItems[0];
            LHS = CreateInspValueFrom(operandItems[1]);
        }

        private IUCIValue CreateInspValueFrom(string valAndType, string conformTo = null)
        {
            if (valAndType.Contains(VALUE_TYPE_SEPARATOR))
            {
                var args = valAndType.Split(new string[] { VALUE_TYPE_SEPARATOR}, StringSplitOptions.None);
                var value = args[0];
                string declaredType = args[1].Equals(string.Empty) ? null : args[1];
                if(conformTo is null)
                {
                    if(declaredType is null)
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

        [TestCase("Dim Hint$\r\nSelect Case Hint$", @"""Here"" To ""Eternity"",""Forever""", "String")] //String
        [TestCase("Dim Hint#\r\nHint#= 1.0\r\nSelect Case Hint#", "10.00 To 30.00, 20.00", "Double")] //Double
        [TestCase("Dim Hint!\r\nHint! = 1.0\r\nSelect Case Hint!", "10.00 To 30.00,20.00", "Single")] //Single
        [TestCase("Dim Hint%\r\nHint% = 1\r\nSelect Case Hint%", "10 To 30,20", "Integer")] //Integer
        [TestCase("Dim Hint&\r\nHint& = 1\r\nSelect Case Hint&", "1000 To 3000,2000", "Long")] //Long
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_uciSelectExprTypeHint(string typeHintExpr, string firstCase, string expected)
        {
            string inputCode =
@"
        Sub Foo()

        <typeHintExprAndSelectCase>
            Case <firstCaseVal>
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<typeHintExprAndSelectCase>", typeHintExpr);
            inputCode = inputCode.Replace("<firstCaseVal>", firstCase);

            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("Dim Hint$\r\nSelect Case x", "Hint$", "String")] //String
        [TestCase("Dim Hint#\r\nHint#= 1.0\r\nSelect Case x", "Hint#", "Double")] //Double
        [TestCase("Dim Hint!\r\nHint! = 1.0\r\nSelect Case x", "Hint!", "Single")] //Single
        [TestCase("Dim Hint%\r\nHint% = 1\r\nSelect Case x", "Hint%", "Integer")] //Integer
        [TestCase("Dim Hint&\r\nHint& = 1\r\nSelect Case x", "Hint&", "Long")] //Long
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_uciCaseClauseTypeHint(string typeHintExpr, string firstCase, string expected)
        {
            string inputCode =
@"
        Sub Foo(x As Variant)

        <typeHintExprAndSelectCase>
            Case <firstCaseVal>
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<typeHintExprAndSelectCase>", typeHintExpr);
            inputCode = inputCode.Replace("<firstCaseVal>", firstCase);

            var result = GetCaseClauseType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("Not x", "x As Long", "Boolean")]
        [TestCase("x", "x As Long", "Long")]
        [TestCase("x < 5", "x As Long", "Boolean")]
        [TestCase("ToLong(True) * .0035", "x As Byte", "Double")]
        [TestCase("True", "x As Byte", "Boolean")]
        [TestCase("ToString(45)", "x As Byte", "String")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_uciSelectExpressionType(string selectExpr, string argList, string expected)
        {
            string inputCode =
@"
        Private Function ToLong(val As Variant) As Long
            ToLong = 5
        End Function

        Private Function ToString(val As Variant) As String
            ToString = ""Foo""
        End Function

        Sub Foo(<argList>)

            Select Case <selectExpr>
                Case 45
                'OK
                Case Else
                'OK
            End Select

        End Sub";

            inputCode = inputCode.Replace("<selectExpr>", selectExpr);
            inputCode = inputCode.Replace("<argList>", argList);

            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("x < 5","False", "Boolean")]
        [TestCase("ToLong(True) * .0035", "45", "Double")]
        [TestCase("True", "x < 5", "Boolean")]
        [TestCase("1 To 10.0", "55 To 100.0", "Double")]
        [TestCase("ToString(45)",@"""Bar""","String")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_uciCaseClauseType(string rangeExpr1, string rangeExpr2, string expected)
        {
            string inputCode =
@"
        Private Function ToLong(val As Variant) As Long
            ToLong = 5
        End Function

        Private Function ToString(val As Variant) As String
            ToString = ""Foo""
        End Function

        Sub Foo(x As Variant)

            Select Case x
                Case <rangeExpr1>, <rangeExpr2>
                'OK
                Case Else
                'OK
            End Select

        End Sub";

            inputCode = inputCode.Replace("<rangeExpr1>", rangeExpr1);
            inputCode = inputCode.Replace("<rangeExpr2>", rangeExpr2);
            var result = GetCaseClauseType(inputCode);
            Assert.AreEqual(expected, result);
        }

        //TODO: Review all uses of tdo.CaseClauseSummary....Looks like there is 
        //a lot of tests that check math rather than SummaryCoverage build-up

        [TestCase("Is < 100", 100, false)]
        [TestCase("Is < 100.49", 100, false)]
        [TestCase("Is < 100#", 100, false)]
        [TestCase("Is < True", -1, false)]
        [TestCase(@"Is < ""100""", 100, false)]
        [TestCase("Is < toVal", 1000, false)]
        [TestCase("Is <= 100", 100, true)]
        [TestCase("Is <= 100.49", 100, true)]
        [TestCase("Is <= 100#", 100, true)]
        [TestCase("Is <= True", -1, true)]
        [TestCase(@"Is <= ""100""", 100, true)]
        [TestCase("Is <= toVal", 1000, true)]
        [TestCase("Is < 45, Is < 100", 100, false)]
        [TestCase("Is < 100, Is < 45", 100, false)]
        [TestCase("Is <= 45, Is <= 100", 100, true)]
        [TestCase("Is <= 100, Is <= 45", 100, true)]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_SummaryCoverageIsLTClause(string firstCase, long isLTMax, bool isLTE)
        {
            string inputCode =
@"
                Private Const fromVal As Long = 500
                Private Const toVal As Long = 1000

                Sub Foo(z As Long)

                Select Case z
                    Case <firstCase>
                    'OK
                End Select

                End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var iSummaryElements = (IUCIRangeClauseFilterTestSupport<long>)GetTestDataObject(inputCode, Tokens.Long).CasesSummary;
            if(iSummaryElements.TryGetIsLTValue(out long ltVal))
            {
                Assert.AreEqual(isLTMax, ltVal, "IsLT value incorrect");
            }
            else
            {
                Assert.Fail("No IsLT value provided");
            }
            if (isLTE)
            {
                Assert.IsTrue(iSummaryElements.SingleValues.Contains(isLTMax), $"SingleValue is missing Value: {isLTMax}");
            }

        }

        [TestCase("Is > 100", 100, false)]
        [TestCase("Is > 100.49", 100, false)]
        [TestCase("Is > 100#", 100, false)]
        [TestCase("Is > True", -1, false)]
        [TestCase(@"Is > ""100""", 100, false)]
        [TestCase("Is > toVal", 1000, false)]
        [TestCase("Is >= 100", 100, true)]
        [TestCase("Is >= 100.49", 100, true)]
        [TestCase("Is >= 100#", 100, true)]
        [TestCase("Is >= True", -1, true)]
        [TestCase(@"Is >= ""100""", 100, true)]
        [TestCase("Is >= toVal", 1000, true)]
        [TestCase("Is > 45, Is > 100", 45, false)]
        [TestCase("Is > 100, Is > 45", 45, false)]
        [TestCase("Is >= 45, Is >= 100", 45, true)]
        [TestCase("Is >= 100, Is >= 45", 45, true)]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_SummaryCoverageIsGTClause(string firstCase, long isGTMin, bool isGTE)
        {
            string inputCode =
@"
        Private Const fromVal As Long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var iSummaryElements = (IUCIRangeClauseFilterTestSupport<long>)GetTestDataObject(inputCode, Tokens.Long).CasesSummary;
            if(iSummaryElements.TryGetIsGTValue(out long gtValue))
            {
                Assert.AreEqual(isGTMin, gtValue, "IsGT value incorrect");
            }
            else
            {
                Assert.Fail("No IsGT provided");
            }
            if (isGTE)
            {
                Assert.AreEqual(true, iSummaryElements.SingleValues.Any(), "SingleValues not updated");
                Assert.AreEqual(true, iSummaryElements.SingleValues.Contains(isGTMin), $"SingleValues does not contain {isGTMin}");
            }
        }

        [TestCase("Is = 100", 100)]
        [TestCase("Is = 100.49", 100)]
        [TestCase("Is = 100#", 100)]
        [TestCase("Is = True", -1)]
        [TestCase(@"Is = ""100""", 100)]
        [TestCase("Is = toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_SummaryCoverageIsEQClause(string firstCase, long isGTMin)
        {
            string inputCode =
@"
        Private Const fromVal As Long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var iSummaryElements = (IUCIRangeClauseFilterTestSupport<long>)GetTestDataObject(inputCode, Tokens.Long).CasesSummary;
            Assert.AreEqual(true, iSummaryElements.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(isGTMin, iSummaryElements.SingleValues.First(), "SingleValue has incorrect Value");

        }

        [TestCase("Is <> 100", 100)]
        [TestCase("Is <> 100.49", 100)]
        [TestCase("Is <> 100#", 100)]
        [TestCase("Is <> True", -1)]
        [TestCase(@"Is <> ""100""", 100)]
        [TestCase("Is <> toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_SummaryCoverageIsNEQClause(string firstCase, long isNEQ)
        {
            string inputCode =
@"
        Private Const fromVal As Long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var iSummaryElements = (IUCIRangeClauseFilterTestSupport<long>)GetTestDataObject(inputCode, Tokens.Long).CasesSummary;
            if (iSummaryElements.TryGetIsGTValue(out long isGT))
            {
                Assert.AreEqual(isNEQ, isGT);
            }
            if (iSummaryElements.TryGetIsLTValue(out long isLT))
            {
                Assert.AreEqual(isNEQ, isLT);
            }
        }

        [TestCase("z < 100", "fromVal < toVal, fromVal = toVal", "Single=-1,0!RelOp=z < 100")]
        [TestCase("100 > z", "fromVal < toVal, fromVal = toVal", "Single=-1,0!RelOp=100 > z")]
        [TestCase("z < 100", "True, True", "Single=-1!RelOp=z < 100")]
        [TestCase("True, True", "z < 100", "Single=-1!RelOp=z < 100")]
        [TestCase("True, False", "z < 100", "Single=-1,0")]
        [TestCase("fromVal < toVal, fromVal = toVal", "z < 100", "Single=-1,0")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_RelationalOpSummaryCoverage(string firstCase, string secondCase, string expected)
        {
            string inputCode =
@"
        Sub Foo(z As Long)

        Private Const fromVal As Long = 500
        Private Const toVal As Long = 1000

        Select Case z
            Case <firstCase>
            'OK
            Case <secondCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            inputCode = inputCode.Replace("<secondCase>", secondCase);
            var itf = GetTestDataObject(inputCode, Tokens.Long).CasesSummary;
            var descriptor = itf.ToString();
            var elements = descriptor.Split(new string[] { "!" }, StringSplitOptions.None);
            var relOps = elements.Any(el => el.StartsWith("RelOp"));

            Assert.AreEqual(expected, itf.ToString());
        }

        [TestCase("IsLT=5", "", "IsLT=5")]
        [TestCase("IsGT=5", "", "IsGT=5")]
        [TestCase("IsLT=5", "IsGT=300", "IsLT=5!IsGT=300")]
        [TestCase("IsLT=5,Range=45:55", "IsGT=300", "IsLT=5!IsGT=300!Range=45:55")]
        [TestCase("IsLT=5,Range=45:55", "IsGT=300,Single=200", "IsLT=5!IsGT=300!Range=45:55!Single=200")]
        [TestCase("IsLT=-2,Range=45:55", "IsGT=300,Single=200,RelOp=x < 50", "IsLT=-2!IsGT=300!Range=45:55!Single=200!RelOp=x < 50")]
        [TestCase("Range=45:55", "Range=60:65", "Range=45:55,60:65")]
        [TestCase("Single=45,Single=46", "Single=60", "Single=45,46,60")]
        [TestCase("RelOp=x < 50", "RelOp=x > 75", "RelOp=x < 50,x > 75")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_ToString(string firstCase, string secondCase, string expectedClauses)
        {
            var sumClauses = TestRangesToSummaryClauses(new string[] { firstCase, secondCase }, Tokens.Long);
            var first = sumClauses[0];
            first.Add(sumClauses[1]);

            Assert.IsTrue(first.ToString().Length > 0, "actual string is zero length");
            Assert.AreEqual(expectedClauses, first.ToString());
        }

        [TestCase("50?Long_To_100?Long", "Long", "Range=50:100")]
        [TestCase("Soup?String_To_Nuts?String", "String", "Range=Nuts:Soup")]
        [TestCase("50.3?Double_To_100.2?Double", "Long", "Range=50:100")]
        [TestCase("50.3?Double_To_100.2?Double", "Double", "Range=50.3:100.2")]
        [TestCase("50_To_100,75_To_125", "Long", "Range=50:100,Range=75:125")]
        [TestCase("50_To_100,175_To_225", "Long", "Range=50:100,Range=175:225")]
        [TestCase("500?Long_To_100?Long", "Long", "Range=100:500")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciAddRangeClauses(string firstCase, string summaryTypeName, string expectedClauses)
        {
            var UUT = RangeClauseFilterFactory.Create(summaryTypeName, ValueFactory);

            var clauses = firstCase.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (var clause in clauses)
            {
                GetBinaryOpValues(clause, out IUCIValue start, out IUCIValue end, out string symbol);
                UUT.AddValueRange(start, end);
            }

            clauses = expectedClauses.Split(new string[] { "," }, StringSplitOptions.None);
            var theVal = CreateTestSummaryCoverage(clauses.ToList(), summaryTypeName);
            Assert.AreEqual(UUT,theVal);
        }

        [TestCase("45", "Long", "Single=45")]
        [TestCase(@"""Foo""", "String", "Single=Foo")]
        [TestCase("45,-500,9", "Long", "Single=45,-500,9")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciAddSingleValue(string firstCase, string summaryTypeName, string expected)
        {
            var UUT = RangeClauseFilterFactory.Create(summaryTypeName, ValueFactory);

            var clauses = firstCase.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (var clause in clauses)
            {
                var theVal = CreateInspValueFrom(clause);
                UUT.AddSingleValue(theVal);
            }
            Assert.IsTrue(UUT.ToString() == expected, $"Actual: {UUT.ToString()} Expected: {expected}");
        }

        [TestCase("x < 100", "Long", "RelOp=x < 100")]
        [TestCase("-1", "Long", "Single=-1")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciAddRelationalOp(string firstCase, string summaryTypeName, string expected)
        {
            var UUT = RangeClauseFilterFactory.Create(summaryTypeName, ValueFactory);

            var clauses = firstCase.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (var clause in clauses)
            {
                var theVal = CreateInspValueFrom(clause);
                UUT.AddRelationalOp(theVal);
            }
            Assert.IsTrue(UUT.ToString() == expected, $"Actual: {UUT.ToString()} Expected: {expected}");
        }

        [TestCase("Is_<_50", "Long", "IsLT=50")]
        [TestCase("Is_<_50,Is_<_25", "Long", "IsLT=50")]
        [TestCase("Is_<_50,Is_<_75", "Long", "IsLT=75")]
        [TestCase("Is_<_50,Is_<_75,Is_>_300", "Long", "IsLT=75!IsGT=300")]
        [TestCase("Is_<=_50", "Long", "IsLT=50!Single=50")]
        [TestCase("Is_<=_50,Is_>=_51", "Long", "IsLT=50!IsGT=51!Single=50,51")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_uciAddIsClauses(string firstCase, string summaryTypeName, string expected)
        {
            var UUT = RangeClauseFilterFactory.Create(summaryTypeName, ValueFactory);

            var clauses = firstCase.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (var clause in clauses)
            {
                GetBinaryOpValues(clause, out IUCIValue start, out IUCIValue end, out string symbol);
                UUT.AddIsClause(end, symbol);
            }
            Assert.AreEqual(expected, UUT.ToString());
        }

        [TestCase("IsLT=45,Range=20:70", "IsLT=45", "Range=45:70")]
        [TestCase("IsLT=45,Range=20:70", "Range=10:70", "IsLT=45")]
        [TestCase("IsLT=45,IsGT=105,Range=20:70", "IsLT=45,Single=200", "IsGT=105,Range=45:70,Single=200")]
        [TestCase("IsLT=45,IsGT=205,Range=20:70,Single=200", "IsLT=45,IsGT=205,Range=20:70", "Single=200")]
        [TestCase("Range=60:80", "Range=20:70,Range=65:100", "")]
        [TestCase("Range=60:80", "IsLT=100", "")]
        [TestCase("Range=60:80", "IsGT=1", "")]
        [TestCase("Single=17", "Range=1:4,Range=7:9,Range=15:20", "")]
        [TestCase("Single=17", "IsLT=45", "")]
        [TestCase("Single=17", "IsGT=-45000", "")]
        [TestCase("Single=17,Single=20", "Single=16,Single=17,Single=18,Single=19", "Single=20")]
        [TestCase("Range=101:149", "Range=101:149,Range=1:100", "")]
        [TestCase("RelOp=x < 50", "Single=-1,Single=0", "")]
        [TestCase("RelOp=x < 50", "Single=-1, RelOp=x < 50", "")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_RemovalRangeClauses(string candidateClauseInput, string existingClauseInput, string expectedClauses)
        {
            var sumClauses = TestRangesToSummaryClauses(new string[] { candidateClauseInput, existingClauseInput }, Tokens.Long);
            var clausesToFilter = sumClauses[0];
            var filter = sumClauses[1];

            var filterResults = RangeClauseFilterFactory.Create(Tokens.Long, ValueFactory);

            filterResults = clausesToFilter.FilterUnreachableClauses(filter);
            if(filterResults.HasCoverage)
            {
                var clauses = expectedClauses.Split(new string[] { "," }, StringSplitOptions.None);
                var expected = CreateTestSummaryCoverage(clauses.ToList(), Tokens.Long);
                Assert.AreEqual(expected, filterResults);
            }
            else
            {
                if (!expectedClauses.Equals(""))
                {
                    Assert.Fail("Function fails to return ISummaryCoverage");
                }
            }
        }

        [TestCase("Range=0:10", "Single=50", "Boolean")]
        [TestCase("Range=True:False", "Single=50", "Boolean")]
        [TestCase("Range=False:True", "Single=50", "Boolean")]
        [TestCase("Single=-5000", "Single=False", "Boolean")]
        [TestCase("Single=True", "Single=0", "Boolean")]
        [TestCase("Single=500", "Single=0", "Boolean")]
        [TestCase("IsLT=5", "IsGT=-5000","Long")]
        [TestCase("IsLT=40,IsGT=40", "Range=35:45", "Long")]
        [TestCase("IsLT=40,IsGT=44", "Range=35:45", "Long")]
        [TestCase("IsLT=40,IsGT=40", "Single=40", "Long")]
        [TestCase("IsGT=240,Range=150:239", "Single=240, Single=0,Single=1,Range=2:150", "Byte")]
        [TestCase("Range=151:255", "Single=150, Single=0,Single=1,Range=2:149", "Byte")]
        [TestCase("IsLT=13,IsGT=30,Range=30:100", "Single=13,Single=14,Single=15,Single=16,Single=17,Single=18,Range=19:30", "Long")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_FiltersAll(string firstCase, string secondCase, string typeName)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase });
            var summaryCoverage = RangeClauseFilterFactory.Create(typeName, ValueFactory);

            foreach (var id in caseToRanges)
            {
                var newSummary = CreateTestSummaryCoverage(id.Value, typeName);
                var filteredResults = newSummary.FilterUnreachableClauses(summaryCoverage);
                if(filteredResults.HasCoverage)
                {
                    summaryCoverage.Add(filteredResults);
                }
            }
            Assert.IsTrue(summaryCoverage.FiltersAllValues, summaryCoverage.ToString());
        }

        [TestCase("IsLT=True, IsGT=True", "RelOp=x > 6", "Single=False,RelOp=Is > True")]
        [TestCase("IsLT=False, IsGT=False", "RelOp=x > 6", "Single=False,RelOp=Is < False")]
        [TestCase("IsGT=False", "RelOp=x > 6", "Single=False")]
        [TestCase("Single=True, Single=False", "Single=True", "")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_FilterBoolean(string firstCase, string secondCase, string expectedClauses)
        {
            var sumClauses = TestRangesToSummaryClauses(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Boolean);

            var candidateClause = sumClauses[0];
            var filter = sumClauses[1];
            var expected = sumClauses[2];
            var filteredResults = candidateClause.FilterUnreachableClauses(filter);
            Assert.AreEqual(expected, filteredResults);
        }

        [TestCase("Range=3:55", "IsLT=6", "IsLT=6,Range=6:55")]
        [TestCase("Range=3:55", "IsGT=6", "IsGT=6,Range=3:6")]
        [TestCase("IsLT=6", "Range=1:5", "IsLT=6")]
        [TestCase("Single=5,Single=6,Single=7", "IsGT=6", "IsGT=6,Single=5,Single=6")]
        [TestCase("Single=5,Single=6,Single=7", "IsLT=6", "IsLT=6,Single=6,Single=7")]
        [TestCase("IsLT=5,IsGT=75", "Single=85", "IsLT=5,IsGT=75")]
        [TestCase("IsLT=5,IsGT=75", "Single=0", "IsLT=5,IsGT=75")]
        [TestCase("Range=45:85", "Single=50", "Range=45:85")]
        [TestCase("Single=5,Single=6,Single=7,Single=8","Range=6:8", "Range=6:8,Single=5")]
        [TestCase("IsLT=400,Range=15:160","Range=500:505", "IsLT=400,Range=500:505")]
        [TestCase("Range=101:149","Range=15:160", "Range=15:160")]
        [TestCase("Range=101:149","Range=15:148", "Range=15:149")]
        [TestCase("Range=150:250,Range=1:100,Range=101:149","Range=25:249", "Range=1:250")]
        [TestCase("Range=150:250,Range=1:100,Range=-5:-2,Range=101:149","Range=25:249", "Range=-5:-2,Range=1:250")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_AddFiltersInteger(string existing, string toAdd, string expectedClause)
        {
            AddFiltersTestSupport(new string[] { existing, toAdd, expectedClause }, Tokens.Long);
        }

        [TestCase("Range=101.45:149.00007", "Range=101.57:110.63", "Range=101.45:149.00007")]
        [TestCase("Range=101.45:149.0007", "Range=15.67:148.9999", "Range=15.67:149.0007")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_AddFiltersRational(string firstCase, string secondCase, string expectedClauses)
        {
            AddFiltersTestSupport(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Double);
        }

        [TestCase(@"Range=""Alpha"":""Omega""", @"Range=""Nuts"":""Soup""", @"Range=""Alpha"":""Soup""")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_AddFiltersString(string firstCase, string secondCase, string expectedClauses)
        {
            AddFiltersTestSupport(new string[] { firstCase, secondCase, expectedClauses }, Tokens.String);
        }

        /*
         * Indeterminant cases are added as unresolved Relational Ops
         * 
        *************************** Is Clause Boolean Truth Table  *********************
        *                          Select Case Value
        *   Resolved Expression     True    False
        *   ****************************************************************************
        *   Is < True               False   False   <= Always False
        *   Is <= True              True    False   
        *   Is > True               False   True    
        *   Is >= True              True    True    <= Always True
        *   Is = True               True    False
        *   Is <> True              False   True
        *   Is > False              False   False   <= Always False
        *   Is >= False             False   True
        *   Is < False              True    False
        *   Is <= False             True    True    <= Always True
        *   Is = False              False   True
        *   Is <> False             True    False
        */

        [TestCase(@"Range=""True:True""", "Single=True", "Single=True")]
        [TestCase(@"Range=""True:False""", "Single=True", "Single=False,Single=True")]
        [TestCase("IsLT=5", "RelOp=x < 5", "Single=False,RelOp=x < 5")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_AddFiltersBoolean(string firstCase, string secondCase, string expectedClauses)
        {
            AddFiltersTestSupport(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Boolean);
        }

        private void AddFiltersTestSupport(string[] input, string typeName)
        {
            Assert.IsTrue(input.Count() >= 2, "At least two rangeClase input strings are neede for this test");

            var sumClauses = TestRangesToSummaryClauses(input, typeName);

            IUCIRangeClauseFilter summary = RangeClauseFilterFactory.Create(typeName, ValueFactory);
            for(var idx = 0; idx <= sumClauses.Count - 2; idx++)
            {
                summary.Add(sumClauses[idx]);
            }

            var expected = sumClauses[sumClauses.Count-1];
            Assert.AreEqual(expected, summary);
        }

        private List<IUCIRangeClauseFilter> TestRangesToSummaryClauses(string[] input, string typeName)
        {
            var caseToRanges = CasesToRanges(input);
            var sumClauses = new List<IUCIRangeClauseFilter>();
            foreach (var id in caseToRanges)
            {
                var newSummary = CreateTestSummaryCoverage(id.Value, typeName);
                sumClauses.Add(newSummary);
            }
            return sumClauses;
        }

        [TestCase("Single=-1,Single=0", "RelOp=x < 3", "Single=-1,Single=0")]
        [TestCase("Range=-5:15", "RelOp=x < 3", "Range=-5:15")]
        [TestCase("IsLT=1", "RelOp=x < 3", "IsLT=1")]
        [TestCase("IsGT=-2", "RelOp=x < 3", "IsGT=-2")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_CoversTrueFalse(string firstCase, string secondCase, string expectedClauses)
        {
            var sumClauses = TestRangesToSummaryClauses(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Long);
            var firstClause = sumClauses[0];
            var secondClause = sumClauses[1];
            var expected = sumClauses[2];

            firstClause.Add(secondClause);

            Assert.AreEqual(expected.ToString(), firstClause.ToString());
        }

        [TestCase("Long")]
        [TestCase("Integer")]
        [TestCase("Byte")]
        [TestCase("Single")]
        [TestCase("Currency")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_Extents(string typeName)
        {
            var summary = RangeClauseFilterFactory.Create(typeName, ValueFactory);

            if(typeName.Equals(Tokens.Long))
            {
                var check = (IUCIRangeClauseFilterTestSupport<long>)summary;
                CheckExtents(check, UCIRangeClauseFilterFactory.IntegerNumberExtents[typeName].Item1, UCIRangeClauseFilterFactory.IntegerNumberExtents[typeName].Item2);
            }
            else if (typeName.Equals(Tokens.Single))
            {
                var check = (IUCIRangeClauseFilterTestSupport<double>)summary;
                CheckExtents(check, CompareExtents.SINGLEMIN, CompareExtents.SINGLEMAX);
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var check = (IUCIRangeClauseFilterTestSupport<decimal>)summary;
                CheckExtents(check, CompareExtents.CURRENCYMIN, CompareExtents.CURRENCYMAX);
            }
        }

        private void CheckExtents<T>(IUCIRangeClauseFilterTestSupport<T> check, T min, T max) where T: IComparable<T>
        {
            if (check.TryGetIsLTValue(out T ltResult) && check.TryGetIsGTValue(out T gtResult))
            {
                Assert.AreEqual(min.ToString().Substring(0, 8), ltResult.ToString().Substring(0, 8), "LT result failed");
                Assert.AreEqual(max.ToString().Substring(0, 8), gtResult.ToString().Substring(0, 8));
            }
            else
            {
                Assert.Fail($"Extents not applied for typeName = {typeof(T).ToString()}");
            }
        }


        private string GetSelectExpressionType(string inputCode)
        {
            var selectStmtValueResults = GetParseTreeValueResults(inputCode, out VBAParser.SelectCaseStmtContext selectStmt);
            var iSelectStmt = InspectionSelectStmtFactory.Create(selectStmt, selectStmtValueResults);
            //var selectStmtValueResults = GetParseTreeValueResultsEvents(inputCode, out IUnreachableCaseInspectionSelectStmt iSelectStmt);
            return iSelectStmt.EvaluationTypeName;
        }

        private string GetCaseClauseType(string inputCode)
        {
            var valueResults = GetParseTreeValueResults(inputCode, out VBAParser.SelectCaseStmtContext selectStmt);
            var iSelectStmt = InspectionSelectStmtFactory.Create(selectStmt, valueResults);
            //var selectStmtValueResults = GetParseTreeValueResultsEvents(inputCode, out IUnreachableCaseInspectionSelectStmt iSelectStmt);
            return iSelectStmt.EvaluationTypeName;
        }

        private IUCIValueResults GetParseTreeValueResults(string inputCode, out VBAParser.SelectCaseStmtContext selectStmt)
        {
            selectStmt = null;
            IUCIValueResults valueResults = null;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                selectStmt = state.ParseTrees.First().Value.GetDescendent<VBAParser.SelectCaseStmtContext>();
                var visitor = ValueVisitorFactory.Create(state, ValueFactory);

                visitor.OnValueResultCreated += Test_OnValueResultCreated;
                valueResults = selectStmt.Accept(visitor);
            }
            return valueResults;
        }

        private IUCIValueResults GetParseTreeValueResultsEvents(string inputCode, out IUnreachableCaseInspectionSelectStmt selectStmt)
        {
            selectStmt = null;
            IUCIValueResults valueResults = null;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var selectContext = state.ParseTrees.First().Value.GetDescendent<VBAParser.SelectCaseStmtContext>();
                selectStmt = InspectionSelectStmtFactory.Create(selectContext, null);
                var visitor = ValueVisitorFactory.Create(state, ValueFactory);
                visitor.OnValueResultCreated += Test_OnValueResultCreated;
                //visitor.OnValueResultCreated += selectStmt.SelectStmt_OnValueResultCreated;
                //selectStmt.Accept(visitor);
                //valueResults = selectStmt.Accept(visitor);
                selectStmt = InspectionSelectStmtFactory.Create(selectContext, valueResults);
            }
            return valueResults;
        }

        private Dictionary<string, List<string>> CasesToRanges(string[] caseClauses)
        {
            var caseToRanges = new Dictionary<string, List<string>>();
            var idx = 0;
            foreach (var cc in caseClauses)
            {
                idx++;
                caseToRanges.Add($"{idx}{cc}", new List<string>());
                var rgs = cc.Split(new string[] { "," }, StringSplitOptions.None);
                foreach (var rg in rgs)
                {
                    caseToRanges[$"{idx}{cc}"].Add(rg.Trim());
                }
            }
            return caseToRanges;
        }

        private IUCIRangeClauseFilter CreateTestSummaryCoverage(List<string> annotations, string typeName)
        {
            var result = RangeClauseFilterFactory.Create(typeName, ValueFactory);
            var clauseItem = string.Empty;
            foreach (var item in annotations)
            {
                var modifiedString = false;
                if(item.Contains(">=") || item.Contains("<="))
                {
                    clauseItem = item.Contains(">=") ? item.Replace(">=", ">") : item.Replace("<=", "<");
                    modifiedString = true;
                }
                else if(item.Contains(" = "))
                {
                    clauseItem = item.Replace(" = ", " @ ");
                    modifiedString = true;
                }
                else
                {
                    clauseItem = item;
                }
                var element = clauseItem.Trim().Split(new string[] { "=" }, StringSplitOptions.None);
                if (element[0].Equals(string.Empty) || element.Count() < 2)
                {
                    continue;
                }
                var clauseType = element[0];
                var clauseExpression = element[1];
                var values = clauseExpression.Split(new string[] { "," }, StringSplitOptions.None);
                foreach ( var expr in values)
                {
                    if (clauseType.Equals("IsLT"))
                    {
                        //if (modifiedString)
                        //{
                        //    clauseExpression.Replace("<", "<=");
                        //    clauseExpression.Replace(">", ">=");
                        //    clauseExpression.Replace(" @ ", " = ");
                        //}
                        var uciVal = ValueFactory.Create(clauseExpression, typeName);
                        result.AddIsClause(uciVal, CompareTokens.LT);
                    }
                    else if (clauseType.Equals("IsGT"))
                    {
                        //if (modifiedString)
                        //{
                        //    clauseExpression.Replace("<", "<=");
                        //    clauseExpression.Replace(">", ">=");
                        //    clauseExpression.Replace(" @ ", " = ");
                        //}
                        var uciVal = ValueFactory.Create(clauseExpression, typeName);
                        result.AddIsClause(uciVal, CompareTokens.GT);
                    }
                    else if (clauseType.Equals("Range"))
                    {
                        var startEnd = clauseExpression.Split(new string[] { ":" }, StringSplitOptions.None);
                        var uciValStart = ValueFactory.Create(startEnd[0], typeName);
                        var uciValEnd = ValueFactory.Create(startEnd[1], typeName);
                        result.AddValueRange(uciValStart, uciValEnd);
                    }
                    else if (clauseType.Equals("Single"))
                    {
                        var uciVal = ValueFactory.Create(clauseExpression, typeName);
                        result.AddSingleValue(uciVal);
                    }
                    else if (clauseType.Equals("RelOp"))
                    {
                        if (modifiedString)
                        {
                            clauseExpression = clauseExpression.Replace("<", "<=");
                            clauseExpression = clauseExpression.Replace(">", ">=");
                            clauseExpression = clauseExpression.Replace(" @ ", " = ");
                        }
                        var uciVal = ValueFactory.Create(clauseExpression, typeName);
                        result.AddRelationalOp(uciVal);
                    }
                    else
                    {
                        Assert.Fail($"Invalid clauseType ({clauseType}) encountered");
                    }
                }
            }
            return result;
        }

        internal struct UnreachableTestDataObject
        {
            public IUCIRangeClauseFilter SummaryCoverage;
            public QualifiedContext<ParserRuleContext> QualifiedContext;
            public VBAParser.SelectCaseStmtContext SelectCaseStmtCtxt;
            public IUnreachableCaseInspectionSelectStmt InspSelectStmt;
            public IUCIRangeClauseFilter CasesSummary;
        }

        private UnreachableTestDataObject GetTestDataObject(string inputCode, string evaluationTypeName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            var tdo = new UnreachableTestDataObject();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                tdo.SelectCaseStmtCtxt = state.ParseTrees.First().Value.GetDescendent<VBAParser.SelectCaseStmtContext>();
                tdo.QualifiedContext = new QualifiedContext<ParserRuleContext>(new QualifiedModuleName(vbe.Object.VBProjects.First()), tdo.SelectCaseStmtCtxt);
                tdo.SummaryCoverage = RangeClauseFilterFactory.Create(evaluationTypeName, ValueFactory);
                tdo.CasesSummary = RangeClauseFilterFactory.Create(tdo.SummaryCoverage.TypeName, ValueFactory);

                var visitor = ValueVisitorFactory.Create(state, ValueFactory);
                var contextResults = new UCIValueResults();
                visitor.OnValueResultCreated += contextResults.OnNewValueResult;
                tdo.SelectCaseStmtCtxt.Accept(visitor);

                tdo.InspSelectStmt = InspectionSelectStmtFactory.Create(tdo.SelectCaseStmtCtxt, contextResults);

                foreach (var caseClause in tdo.SelectCaseStmtCtxt.caseClause())
                {
                    var summary = RangeClauseFilterFactory.Create(tdo.SummaryCoverage.TypeName, ValueFactory);
                    foreach (var range in caseClause.rangeClause())
                    {
                        var inspR = InspectionRangeFactory.Create(tdo.SummaryCoverage.TypeName, range, contextResults);
                        summary.Add(inspR.AsFilter);
                    }
                    tdo.CasesSummary.Add(summary);
                }
            }
            return tdo;
        }

        #region oldTests
        /**/
        [TestCase("String", @"""Foo""", @"""Bar""")]
        [TestCase("Long", "450000", "850000")]
        [TestCase("Integer", "4500", "8500")]
        [TestCase("Byte", "3", "254")]
        [TestCase("Double", "45000.345", "55000.25")]
        [TestCase("Single", "45.345", "55.25")]
        [TestCase("Currency", "4.34578", "5.25869")]
        [TestCase("Boolean", "True", "False")]
        [TestCase("Boolean", "55", "0")]
        [TestCase("Long", "-450000", "850000")]
        [TestCase("Integer", "-4500", "8500")]
        [TestCase("Double", "-45000.345", "55000.25")]
        [TestCase("Single", "-45.345", "55.25")]
        [TestCase("Currency", "-4.34578", "5.25869")]
        [TestCase("Boolean", "-55", "0")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SingleUnreachableAllTypes(string type, string value1, string value2)
        {
            string inputCode =
@"Sub Test(x As <Type>)

        Const firstVal As <Type> = <Value1>
        Const secondVal As <Type> = <Value2>

        Select Case x
            Case firstVal, secondVal
            'OK
            Case firstVal
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<Type>", type);
            inputCode = inputCode.Replace("<Value1>", value1);
            inputCode = inputCode.Replace("<Value2>", value2);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestCase("Long", "2147486648#", "-2147486649#")]
        [TestCase("Integer", "40000", "-50000")]
        [TestCase("Byte", "256", "-1")]
        [TestCase("Currency", "922337203685490.5808", "-922337203685477.5809")]
        [TestCase("Single", "3402824E38", "-3402824E38")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_ExceedsLimits(string type, string value1, string value2)
        {
            string inputCode =
@"Sub Foo(x As <Type>)

        Const firstVal As <Type> = <Value1>
        Const secondVal As <Type> = <Value2>

        Select Case x
            Case firstVal
            'Unreachable
            Case secondVal
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<Type>", type);
            inputCode = inputCode.Replace("<Value1>", value1);
            inputCode = inputCode.Replace("<Value2>", value2);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestCase("x Or x < 5")]
        [TestCase("x = 1 Xor x < 5")]
        [TestCase("x And x < 5")]
        [TestCase("Not x")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LogicalOpSelectCase(string booleanOp)
        {
            string inputCode =
@"Sub Foo(x As Long)
        Select Case <booleanOp>
            Case True
            'OK
            Case False 
            'Unreachable
            Case -5
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<booleanOp>", booleanOp);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_RelationalOpSelectCase()
        {
            string inputCode =
@"Sub Foo(x As Long)

        Private Const fromVal As Long = 500
        Private Const toVal As Long = 1000

        Select Case x
           Case fromVal < toVal
            'OK
           Case x < 100
            'OK
           Case fromVal = toVal , fromVal < toVal
            'OK
            Case x > 300
            'Unreachable
            Case x = 200
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_RelationalOpExpression()
        {
            string inputCode =
@"Sub Foo(x As Long)

        Private Const fromVal As Long = 500
        Private Const toVal As Long = 1000

        Select Case x
           Case toVal < fromVal * 6
            'OK
           Case True
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestCase("Not fromVal", "False")]
        [TestCase("Not toVal", "True")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LogicalOpUnary(string caseClause, string expected)
        {
            string inputCode =
@"Sub Foo(x As Long)

        Private Const fromVal As Long = 500
        Private Const toVal As Long = 0

        Select Case x
           Case <caseClause>
            'OK
           Case True
            'Unreachable
        End Select

        End Sub";

            inputCode = inputCode.Replace("<caseClause>", caseClause);
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var visitorFactory = FactoriesFactoryTest.CreateIUCIParseTreeValueVisitorFactory();
                var ptVisitor = visitorFactory.Create(state, ValueFactory);
                var selectStmtContext = state.ParseTrees.First().Value.GetDescendent<VBAParser.SelectCaseStmtContext>();
                //TODO: we need this in many places
                ptVisitor.OnValueResultCreated += Test_OnValueResultCreated;
                var result = selectStmtContext.Accept(ptVisitor);
                var logicalNotContext = selectStmtContext.GetDescendent<VBAParser.LogicalNotOpContext>();
                Assert.AreEqual(expected, result.GetValueText(logicalNotContext));
            }
        }

        [TestCase("Is > 8", "12", "9")]
        [TestCase("Is >= 8", "12", "8")]
        [TestCase("Is < 8", "-56", "7")]
        [TestCase("Is <= 8", "-56", "8")]
        [TestCase("Is <> 8", "-56", "5000")]
        [TestCase("Is = 8", "16 / 2", "4 * 2")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IsStmt(string isStmt, string unreachableValue1, string unreachableValue2)
        {
            string inputCode =
@"Sub Foo(z As Long)

        Select Case z
            Case <IsStmt>
            'OK
            Case <Unreachable1>
            'Unreachable
            Case <Unreachable2>
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<IsStmt>", isStmt);
            inputCode = inputCode.Replace("<Unreachable1>", unreachableValue1);
            inputCode = inputCode.Replace("<Unreachable2>", unreachableValue2);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 2 = 49, x ^ 3 = 8")]
        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30")]
        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "(CLng(VBA.Rnd() * 100) * x) < 30, x ^ 2 = 49")]
        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 3 = 8")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CopyPaste(string complexClause1, string complexClause2)
        {
            string inputCode =
@"Sub Foo(x As Long)

                Select Case x
                    Case <complexClause1>
                    'OK
                    Case <complexClause2>
                    'Unreachable - detected by text compare of range clause(s)
                End Select

                End Sub";
            inputCode = inputCode.Replace("<complexClause1>", complexClause1);
            inputCode = inputCode.Replace("<complexClause2>", complexClause2);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestCase("Long", "5000 - 1000", "4000")]
        [TestCase("Double", "50.00 - 10.00", "40.00")]
        [TestCase("Currency", "50.00 - 10.00", "40.00")]
        [TestCase("Single", "50.00 - 10.00", "40.00")]
        [TestCase("Long", "5000 + 1000", "6000")]
        [TestCase("Double", "50.00 + 10.00", "60.00")]
        [TestCase("Single", "50.00 + 10.00", "60.00")]
        [TestCase("Long", "50 * 10", "500")]
        [TestCase("Double", "50.00 * 10.00", "500.00")]
        [TestCase("Single", "50.00 * 10.00", "500.00")]
        [TestCase("Long", "5000 / 1000", "5")]
        [TestCase("Double", "50.00 / 10.00", "5.0")]
        [TestCase("Currency", "50.00 / 10.00", "5.0")]
        [TestCase("Single", "50.00 / 10.00", "5.0")]
        [TestCase("Single", "52.00 Mod 10.00", "2.0")]
        [TestCase("Single", "2.00 ^ 3.00", "8.0")]
        [TestCase("Integer", "58 Mod 4", "2")]
        [TestCase("Integer", "2 ^ 3", "8")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseClauseHasBinaryMathOp(string type, string mathOp, string unreachable)
        {
            string inputCode =
@"
        Sub Foo(z As <Type>)

        Select Case z
            Case <mathOp>
            'OK
            Case <unreachable>
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<Type>", type);
            inputCode = inputCode.Replace("<mathOp>", mathOp);
            inputCode = inputCode.Replace("<unreachable>", unreachable);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Ignore("Invalid")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_PowOpEvaluationAlgebraNoDetection()
        {
            const string inputCode =
@"Sub Foo(x As Long)

        Select Case x
            Case x ^ 2 = 49
            'OK
            Case x = 7
            'Unreachable, but not detected - math/algebra on the Select Case variable yet to be supported
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_NumberRangeConstants()
        {
            const string inputCode =
@"Sub Foo(x As Long, z As Double)

        Const JAN As Long = 1
        Const DEC As Long = 12
        Const AUG As Long = 8

        Select Case z * x
            Case JAN To DEC
            'OK
            Case AUG
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestCase(@"1 To ""Forever""", 1, 1)]
        [TestCase(@"""Fifty-Five"" To 1000", 1, 1)]
        [TestCase("z To 1000", 1, 1)]
        [TestCase("50 To z", 1, 1)]
        [TestCase(@"z To 1000, 95, ""TEST""", 1, 0)]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_NumberRangeMixedTypes(string firstCase, int unreachableCount, int mismatchCount)
        {
            string inputCode =
@"Sub Foo(x As Long, z As String)

        Select Case x
            Case <firstCase>
            'Mismatch - unreachable
            Case 1 To 50
            'OK
            Case 45
            'Unreachable
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            CheckActualResultsEqualsExpected(inputCode, unreachable: unreachableCount, mismatch: mismatchCount);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseElseIsClausePlusRange()
        {
            const string inputCode =
@"Sub Foo(x as Long)

        Select Case x
            Case Is > 200
            'OK
            Case 50 To 200
            'OK
            Case Is < 50
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseElseIsClausePlusRangeAndSingles()
        {
            const string inputCode =
@"Sub Foo(x as Long)

        Select Case x
            Case 53,54
            'OK
            Case Is > 200
            'OK
            Case 55 To 200
            'OK
            Case Is < 50
            'OK
            Case 50,51,52
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_NestedSelectCase()
        {
            const string inputCode =
@"Sub Foo(x As Long, z As Long) 

        Select Case x
            Case 1 To 10
            'OK
            Case 9
            'Unreachable
            Case 11
            Select Case  z
                Case 5 To 25
                'OK
                Case 6
                'Unreachable
                Case 8
                'Unreachable
                Case 15
                'Unreachable
            End Select
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 4);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_NestedSelectCases()
        {
            const string inputCode =
@"Sub Foo(x As String, z As String )

        Select Case x
            Case ""Foo"", ""Bar"", ""Goo""
            'OK
            Case ""Foo""
            'Unreachable
            Case ""Food""
                Select Case  z
                    Case ""Food"", ""Bard"",""Good""
                    'OK
                    Case ""Bar""
                    'OK
                    Case ""Foo""
                    'OK
                    Case ""Goo""
                    'OK
                End Select
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_NestedSelectCaseSUnreachable()
        {
            const string inputCode =
@"Sub Foo(x As String, z As String)

Select Case x
            Case ""Foo"", ""Bar""
            'OK
            Case ""Foo""
            'Unreachable
            Case ""Food""
            Select Case  z
                Case ""Bar"",""Goo""
                'OK
                Case ""Bar""
                'Unreachable
                Case ""Foo""
                'OK
                Case ""Goo""
                'Unreachable
            End Select
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SimpleLongCollisionConstantEvaluation()
        {
            const string inputCode =
@"

        private const BASE As Long = 10
        private const MAX As Long = BASE ^ 2

        Sub Foo(x As Long)

        Select Case x
            Case 100
            'OK
            Case MAX 
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }
        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_MixedSelectCaseTypes()
        {
            const string inputCode =
@"

        private const MAXValue As Long = 5
        private const TwentyFiveCents As Double = .25
        private const MINCoins As Long = 4

        Sub Foo(numQuarters As Byte)

        Select Case numQuarters * TwentyFiveCents
            Case 1.25 To 10.00
            'OK
            Case MAXValue 
            'Unreachable
            Case MINCoins * TwentyFiveCents
            'OK
            Case MINCoins * 2
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_ExceedsIntegerButIncludesAccessibleValues()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

        Select Case x
            Case -50000
            'Exceeds Integer values and unreachable
            Case 10,11,12
            'OK
            Case 15, 40000
            'Exceeds Integer value - but other value makes case reachable....no Error
            Case Is < 4
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IntegerWithDoubleValue()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

        Select Case x
            Case Is < -50
            'OK
            Case 214.0
            'OK - ish
            Case -214#
            'unreachable
            Case 98
            'OK
            Case 5 To 25, 50, 80
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_VariantSelectCase()
        {
            const string inputCode =
@"Sub Foo( x As Variant)

        Select Case x
            Case .4 To .9
            'OK
            Case 0.23
            'OK
            Case 0.55
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_VariantSelectCaseInferFromConstant()
        {
            const string inputCode =
@"Sub Foo( x As Variant)

        private Const TheValue As Double = 45.678
        private Const TheUnreachableValue As Long = 25

        Select Case x
            Case TheValue * 2
            'OK
            Case 0 To TheValue
            'OK
            Case TheUnreachableValue
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_VariantSelectCaseInferFromConstant2()
        {
            const string inputCode =
@"Sub Foo( x As Variant)

        private Const TheValue As Double = 45.678
        private Const TheUnreachableValue As Long = 77

        Select Case x
            Case Is > TheValue
            'OK
            Case 0 To TheValue - 20
            'OK
            Case TheUnreachableValue
            'Unreachable
            Case 55
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_BuiltInSelectCase()
        {
            const string inputCode =
@"
Function Random() As Double
    Random = VBA.Rnd()
End Function

Sub Foo( x As Variant)

        Select Case Random()
            Case .4 To .9
            'OK
            Case 0.23
            'OK
            Case 0.55
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

         //The test cases below represent the truth table
        //for Is clauses for boolean Select Case Statements.
        //See UCIRangeClauseFilter.AddIsClauseBoolean(...)
        //Cases that always resolve to True (or False) are stored as Single values.
        //All others are stored as variable RelationalOp expressions
        [TestCase("Is < True", "Single=False")]
        [TestCase("Is <= True", "RelOp=Is <= True")]
        [TestCase("Is > True", "RelOp=Is > True")]
        [TestCase("Is >= True", "Single=True")]
        [TestCase("Is = True", "RelOp=Is = True")]
        [TestCase("Is <> True", "RelOp=Is <> True")]
        [TestCase("Is > False", "Single=False")]
        [TestCase("Is >= False", "RelOp=Is >= False")]
        [TestCase("Is < False", "RelOp=Is < False")]
        [TestCase("Is <= False", "Single=True")]
        [TestCase("Is = False", "RelOp=Is = False")]
        [TestCase("Is <> False", "RelOp=Is <> False")]
        [Category("Inspections")]
        public void UnreachableCaseInspUnit_BooleanIsClauseTruthTable(string firstCase, string expected)
        {
            string inputCode =
@"Sub Foo( x As Boolean)

        Select Case x
            Case <firstCase>
            'OK
            Case Else
            'Foo
        End Select
End Sub";
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var results = GetParseTreeValueResults(inputCode, out VBAParser.SelectCaseStmtContext selectStmtContext);
            var range = selectStmtContext.GetDescendent<VBAParser.RangeClauseContext>();
            var inspRange = InspectionRangeFactory.Create(range, results);
            inspRange.EvaluationTypeName = Tokens.Boolean;

            var expectedFilters = TestRangesToSummaryClauses(new string[] { expected }, Tokens.Boolean);
            Assert.AreEqual(expectedFilters.First().ToString(), inspRange.AsFilter.ToString());
        }

        [TestCase("True", "Is <= False", 2)]
        [TestCase("Is >= True", "False", 2)]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_BooleanRelationalOps(string firstCase, string secondCase, int expected)
        {
            string inputCode =
@"Sub Foo( x As Boolean)

        Select Case x
            Case <firstCase>
            'OK
            Case <secondCase>
            'Unreachable
            Case 95
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            inputCode = inputCode.Replace("<secondCase>", secondCase);
            CheckActualResultsEqualsExpected(inputCode, unreachable: expected);
        }

        [Ignore("Invalid")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LongCollisionIndeterminateCase()
        {
            const string inputCode =
@"Sub Foo( x As Long, y As Double)

        Select Case x
            Case x > -3000
            'OK
            Case x < y
            'OK - indeterminant
            Case 95
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Ignore("this is valid, but do we want to support inspection for this?")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LongCollisionMultipleVariablesSameType()
        {
            const string inputCode =
@"Sub Foo( x As Long, y As Long)

        Select Case x * y
            Case x > -3000
            'OK
            Case y > -3000
            'OK
            Case x < y
            'OK - indeterminant
            Case 95
            'OK - this gives a false positive when evaluated as if 'x' or 'y' is the only select case variable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [Ignore("this is valid, but do we want to support inspection for this?")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LongCollisionMultipleVariablesDifferentType()
        {
            const string inputCode =
@"Sub Foo( x As Long, y As Double)

        Select Case x * y
            Case x > -3000
            'OK
            Case y > -3000
            'OK
            Case x < y
            'OK - indeterminant
            Case 95
            'OK
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Ignore("this is valid, but do we want to support inspection for this?")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SelectExpressionMathop()
        {
            const string inputCode =
@"Sub Foo( x As Long, y As Double)

        Select Case x * y
            Case x > -3000
            'OK
            Case y > -3000
            'OK
            Case x < y
            'OK - indeterminant
            Case Is > 5
            'OK
            Case 95
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [Ignore("this is valid, but do we want to support inspection for this?")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LongCollisionVariableAndConstantDifferentType()
        {
            const string inputCode =
@"Sub Foo( x As Long)

        private const y As Double = 0.5

        Select Case x * y
            Case x > -3000
            'OK
            Case y > -3000
            'Unreachable
            Case x < y
            'OK - indeterminant
            Case 95
            'OK - this gives a false positive when evaluated as if 'x' is the only select case variable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [Ignore("this is valid, but do we want to support inspection for this?")]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_LongCollisionUnaryMathOperation()
        {
            const string inputCode =
@"Sub Foo( x As Long, y As Double)

        Select Case -x
            Case x > -3000
            'OK
            Case y > -3000
            'Cannot disqualify other, or be disqualified, except by another y > ** statement
            Case x < y
            'OK - indeterminant
            Case 95
            'unreachable - not evaluated
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_BooleanExpressionUnreachableCaseElseInvertBooleanRange()
        {
            const string inputCode =
@"
        Private Function Random() As Double
            Random = VBA.Rnd()
        End Function


        Sub Foo(x As Boolean)


        Select Case Random() > 0.5
            Case False To True 
            'OK
            Case True
            'Unreachable
            Case Else
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_StringWhereLongShouldBe()
        {
            const string inputCode =
@"Sub Foo(x As Long)

        Select Case x
            Case 1 To 49
            'OK
            Case 50
            'OK
            Case ""Test""
            'Unreachable
            Case ""85""
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_MixedTypes()
        {
            const string inputCode =
@"Sub Foo(x As Long)

        Select Case x
            Case 1 To 49
            'OK
            Case ""Test"", 100, ""92""
            'OK - ""Test"" will not be evaluated
            Case ""85""
            'OK
            Case 2
            'Unreachable
            Case 92
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_StringWhereLongShouldBeIncludeLongAsString()
        {
            const string inputCode =
@"Sub Foo(x As Long)

        Select Case x
            Case 1 To 49
            'OK
            Case ""51""
            'OK
            Case ""Hello World""
            'Unreachable
            Case 50
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CascadingIsStatements()
        {
            const string inputCode =
@"Sub Foo(LNumber As Long)

        Select Case LNumber
            Case Is < 100
                'OK
            Case Is < 200
                'OK
            Case Is < 300
                'OK
            Case Else
                'OK
            End Select
        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CascadingIsStatementsGT()
        {
            const string inputCode =
@"Sub Foo(LNumber As Long)

        Select Case LNumber
            Case Is > 300
            'OK
            Case Is > 200
            'OK  
            Case Is > 100
            'OK  
            Case Else
            'OK
            End Select
        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IsStatementUnreachableGT()
        {
            const string inputCode =
@"Sub Foo(x As Long)

        Select Case x
            Case Is > 100
                'OK  
            Case Is > 200
                'unreachable  
            Case Is > 300
                'unreachable
            Case Else
                'OK
            End Select
        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IsStatementUnreachableLT()
        {
            const string inputCode =
@"Sub Foo(x As Long)

        Select Case x
            Case Is < 300
                'OK  
            Case Is < 200
                'unreachable  
            Case Is < 100
                'unreachable
            Case Else
                'OK
            End Select
        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IsStmtToIsStmtCaseElseUnreachableUsingIs()
        {
            const string inputCode =
@"Sub Foo(z As Long)

        Select Case z
            Case Is <> 5 
            'OK
            Case Is = 5
            'OK
            Case 400
            'Unreachable
            Case Else
            'Unreachable
        End Select
        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseClauseHasParens()
        {
            const string inputCode =
@"
        Sub Foo(z As Long)

        private const maxValue As Long = 5000
        private const subtract As Long = 2000

        Select Case z
            Case (maxValue - subtract) * 10
            'OK
            Case 30000
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseClauseHasMultipleParens()
        {
            const string inputCode =
@"
        Sub Foo(z As Long)

        private const maxValue As Long = 5000
        private const subtractValue As Long = 2000

        Select Case z
            Case (maxValue - subtractValue) * (55 - 35) / 10
            'OK
            Case 6000
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SelectCaseHasMultOpWithFunction()
        {
            const string inputCode =
@"
        Function Bar() As Long
            Bar = 5
        End Function

        Sub Foo(z As Long)

        Select Case Bar() * z
            Case Is > 5000
            'OK
            Case 5000
            'OK
            Case 5001
            'Unreachable
            Case 10000
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseClauseHasMultOpInParens()
        {
            const string inputCode =
@"
        Sub Foo(z As Long)

        private const maxValue As Long = 5000

        Select Case (((z)))
            Case ((2 * maxValue))
            'OK
            Case 10000
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseClauseHasMultOp2Constants()
        {
            const string inputCode =
@"
        Sub Foo(z As Long)

        private const maxValue As Long = 5000
        private const minMultiplier As Long = 2

        Select Case z
            Case maxValue / minMultiplier
            'OK
            Case 2500
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_EnumerationNumberRangeNoDetection()
        {
            const string inputCode =
@"
        private Enum Weekday
            Sunday = 1
            Monday = 2
            Tuesday = 3
            Wednesday = 4
            Thursday = 5
            Friday = 6
            Saturday = 7
            End Enum

        Sub Foo(z As Weekday)

        Select Case z
            Case Weekday.Monday To Weekday.Saturday
            'OK
            Case z = Weekday.Tuesday
            'OK
            Case Weekday.Wednesday
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_EnumerationNumberRangeNonConstant()
        {
            const string inputCode =
@"
        private Enum BitCountMaxValues
            max1Bit = 2 ^ 0
            max2Bits = 2 ^ 1 + max1Bit
            max3Bits = 2 ^ 2 + max2Bits
            max4Bits = 2 ^ 3 + max3Bits
        End Enum

        Sub Foo(z As BitCountMaxValues)

        Select Case z
            Case 7
            'OK
            Case BitCountMaxValues.max3Bits
            'Unreachable
            Case BitCountMaxValues.max4Bits
            'OK
            Case 15
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_EnumerationLongCollision()
        {
            const string inputCode =
@"
        private Enum BitCountMaxValues
            max1Bit = 2 ^ 0
            max2Bits = 2 ^ 1 + max1Bit
            max3Bits = 2 ^ 2 + max2Bits
            max4Bits = 2 ^ 3 + max3Bits
        End Enum

        Sub Foo(z As BitCountMaxValues)

        Select Case z
            Case BitCountMaxValues.max3Bits
            'OK
            Case 7
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_EnumerationNumberRangeConflicts()
        {
            const string inputCode =
@"
                private Enum Fruit
                    Apple = 10
                    Pear = 20
                    Orange = 30
                    End Enum

                Sub Foo(z As Fruit)

                Select Case z
                    Case Apple
                    'OK
                    Case Pear 
                    'OK     
                    Case Orange        
                    'OK
                    Case Else
                    'OK - avoid flagging CaseElse for enums so guard clauses such as below are retained
                    Err.Raise 5, ""MyFunction"", ""Invalid value given for the enumeration.""
                End Select

                End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0, caseElse: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_EnumerationNumberCaseElse()
        {
            const string inputCode =
@"
                private Enum Fruit
                    Apple = 10
                    Pear = 20
                    Orange = 30
                    End Enum

                Sub Foo(z As Fruit)

                Select Case z
                    Case Is <> Apple
                    'OK
                    Case Apple 
                    'OK     
                    Case Else
                    'unreachable - Guard clause will always be skipped
                    Err.Raise 5, ""MyFunction"", ""Invalid value given for the enumeration.""
                End Select

                End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_CaseElseByte()
        {
            const string inputCode =
@"
        Sub Foo(z As Byte)

        Select Case z
            Case Is >= 2
            'OK
            Case 0,1
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [TestCase("( z * 3 ) - 2", "Is > maxValue", 2)]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SelectCaseUsesConstantReferenceExpr(string selectExpr, string firstCase, int expected)
        {
            string inputCode =
@"
        private Const maxValue As Long = 5000

        Sub Foo(z As Long)

        Select Case <selectExpr>
            Case <firstCase>
            'OK
            Case 15
            'OK
            Case 6000
            'Unreachable
            Case 8500
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<selectExpr>", selectExpr);
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            CheckActualResultsEqualsExpected(inputCode, unreachable: expected);
        }

        //TODO: Still a relevant test? (after it passes)
        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SelectCaseUsesConstantIndeterminantExpression()
        {
            const string inputCode =
@"
        private Const maxValue As Long = 5000

        Sub Foo(z As Long)

        Select Case z
            Case z > maxValue / 2
            'OK
            Case z > maxValue
            'OK
            Case 15
            'OK
            Case 8500
            'OK
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SelectCaseIsFunction()
        {
            const string inputCode =
@"
        Function Bar() As Long
            Bar = 5
        End Function

        Sub Foo()

        Select Case Bar()
            Case Is > 5000
            'OK
            Case 5000
            'OK
            Case 5001
            'Unreachable
            Case 10000
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_SelectCaseIsFunctionWithParams()
        {
            const string inputCode =
@"
        Function Bar(x As Long, y As Double) As Long
            Bar = 5
        End Function

        Sub Foo(firstVar As Long, secondVar As Double)

        Select Case Bar( firstVar, secondVar )
            Case Is > 5000
            'OK
            Case 5000
            'OK
            Case 5001
            'Unreachable
            Case 10000
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IsStmtAndNegativeRange()
        {
            const string inputCode =
@"Sub Foo(z As Long)

        Select Case z
            Case Is < 8
            'OK
            Case -10 To -3
            'Unreachable
            Case 0
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspFunctional_IsStmtAndNegativeRangeWithConstants()
        {
            const string inputCode =
@"
        private const START As Long = 10
        private const FINISH As Long = 3

        Sub Foo(z As Long)
        Select Case z
            Case Is < 8
            'OK
            Case -(START * 4) To -(FINISH * 2) 
            'Unreachable
            Case 0
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }
/**/
#endregion
        private void CheckActualResultsEqualsExpected(string inputCode, int unreachable = 0, int mismatch = 0, int caseElse = 0)
        {
            var expected = new Dictionary<string, int>
            {
                { InspectionsUI.UnreachableCaseInspection_Unreachable, unreachable },
                { InspectionsUI.UnreachableCaseInspection_TypeMismatch, mismatch },
                { InspectionsUI.UnreachableCaseInspection_CaseElse, caseElse },
            };

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            IEnumerable<Rubberduck.Parsing.Inspections.Abstract.IInspectionResult> actualResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnreachableCaseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }
            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(InspectionsUI.UnreachableCaseInspection_Unreachable));
            var actualMismatches = actualResults.Where(ar => ar.Description.Equals(InspectionsUI.UnreachableCaseInspection_TypeMismatch));
            var actualUnreachableCaseElses = actualResults.Where(ar => ar.Description.Equals(InspectionsUI.UnreachableCaseInspection_CaseElse));

            var actualMsg = BuildResultString(actualUnreachable.Count(), actualMismatches.Count(), actualUnreachableCaseElses.Count());
            var expectedMsg = BuildResultString(expected[InspectionsUI.UnreachableCaseInspection_Unreachable], expected[InspectionsUI.UnreachableCaseInspection_TypeMismatch], expected[InspectionsUI.UnreachableCaseInspection_CaseElse]);

            Assert.AreEqual(expectedMsg, actualMsg);
        }
        private string BuildResultString(int unreachableCount, int mismatchCount, int caseElseCount)
        {
            return  $"Unreachable={unreachableCount}, Mismatch={mismatchCount}, CaseElse={caseElseCount}";
        }
    }
}
