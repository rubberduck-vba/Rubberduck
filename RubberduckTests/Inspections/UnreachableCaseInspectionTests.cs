using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
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

        [TestCase(@"""105""", @"""105""")]
        [TestCase("105", "105")]
        [TestCase("105.6", "105.6")]
        [TestCase("45.2", "45.2")]
        [TestCase("True", "-1")]
        [TestCase("False", "0")]
        [TestCase("32.000023@", "32.000023")]
        [TestCase("32.000023!", "32.000023")]
        [TestCase("32.000023#", "32.000023")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ParseTreeValueConversionTests(string testValue, string checkValue)
        {
            var ctxtValue = new ParseTreeValue(testValue);
            var convertible = checkValue.Replace("\"", "");

            var testDouble = Convert.ToDouble(convertible);
            Assert.AreEqual(testDouble, ctxtValue.AsDouble(),  "Double Failed");

            var testDecimal = Convert.ToDecimal(convertible);
            Assert.AreEqual(testDecimal, ctxtValue.AsCurrency(), "Decimal Failed");

            var testLong = Convert.ToInt64(testDouble);
            Assert.AreEqual(testLong, ctxtValue.AsLong(), "Long Failed");

            var testInt = Convert.ToInt32(testLong);
            Assert.AreEqual(testInt, ctxtValue.AsInt(), "Integer Failed");

            if (testLong > 0 && testLong < 256)
            {
                var testByte = Convert.ToByte(testLong);
                Assert.AreEqual(testByte, ctxtValue.AsByte(), "Byte Failed");
            }

            var testBool = Convert.ToBoolean(testLong);
            Assert.AreEqual(testBool, ctxtValue.AsBoolean(), "Boolean Failed");

            if(testValue.Equals(Tokens.True) || testValue.Equals(Tokens.False))
            {
                Assert.AreEqual(testValue, ctxtValue.AsString(), "String Failed");
            }
            else
            {
                Assert.AreEqual(checkValue, ctxtValue.AsString(), "String Failed");
            }
        }

        [TestCase("What@", "What@")]
        [TestCase("What!", "What!")]
        [TestCase("What#", "What#")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NonNumberConversions(string firstCase, string value)
        {
            var ctxtValue = new ParseTreeValue(firstCase);
            Assert.IsFalse(ctxtValue.AsLong().HasValue);
            Assert.AreEqual(ctxtValue.AsString(),value);
        }

        [TestCase("10.5", "105.6", "Long")]
        [TestCase("10.5", "105.6", "Integer")]
        [TestCase("10.5", "105.6", "Double")]
        [TestCase("10.5", "105.6", "Byte")]
        [TestCase("-1", "-0", "Boolean")]
        [TestCase("Apples", "Oranges", "String")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ParseTreeValueOperatorTests(string smallVal, string bigVal, string typeName)
        {
            var smallValue = new ParseTreeValue(smallVal, typeName);
            var bigValue = new ParseTreeValue(bigVal, typeName);

            Assert.True(smallValue < bigValue, $"{typeName}: LT Failed");
            Assert.True(smallValue <= bigValue, $"{typeName}: LTE Failed");
            Assert.True(bigValue > smallValue, $"{typeName}: GT Failed");
            Assert.True(bigValue >= smallValue, $"{typeName}: GTE Failed");
            Assert.False(bigValue == smallValue, $"{typeName}: EQ Failed");
            Assert.True(bigValue != smallValue, $"{typeName}: NEQ Failed");
        }

        [TestCase("10_*_2", "20", "Long")]
        [TestCase("10_/_2", "5", "Long")]
        [TestCase("10_+_2", "12", "Long")]
        [TestCase("10_-_2", "8", "Long")]
        [TestCase("10_Pow_2", "100", "Long")]
        [TestCase("10_Mod_2", "0", "Long")]
        [TestCase("10_*_2", "20", "Double")]
        [TestCase("10_/_2", "5", "Double")]
        [TestCase("10_+_2", "12", "Double")]
        [TestCase("10_-_2", "8", "Double")]
        [TestCase("10_Pow_2", "100", "Double")]
        [TestCase("10_Mod_2", "0", "Double")]
        [TestCase("10_*_2", "20", "Byte")]
        [TestCase("10_/_2", "5", "Byte")]
        [TestCase("10_+_2", "12", "Byte")]
        [TestCase("10_-_2", "8", "Byte")]
        [TestCase("10_Pow_2", "100", "Byte")]
        [TestCase("10_Mod_2", "0", "Byte")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ParseTreeValueMathOperatorTests(string operands, string result, string typeName)
        {
            var separator = new string[] { "_" };
            var lhs = operands.Split(separator, StringSplitOptions.None)[0];
            var op = operands.Split(separator, StringSplitOptions.None)[1];
            var rhs = operands.Split(separator, StringSplitOptions.None)[2];
            var LHS = new ParseTreeValue(lhs, typeName);
            var RHS = new ParseTreeValue(rhs, typeName);

            if (op.Equals("*"))
            {
                var testResult = LHS * RHS;
                Assert.AreEqual(testResult, new ParseTreeValue(result, typeName), $"{typeName}: '{op}' operator Failed");
            }
            else if (op.Equals("/"))
            {
                var testResult = LHS / RHS;
                Assert.AreEqual(testResult, new ParseTreeValue(result, typeName), $"{typeName}: '{op}' operator Failed");
            }
            else if (op.Equals("+"))
            {
                var testResult = LHS + RHS;
                Assert.AreEqual(testResult, new ParseTreeValue(result, typeName), $"{typeName}: '{op}' operator Failed");
            }
            else if (op.Equals("-"))
            {
                var testResult = LHS - RHS;
                Assert.AreEqual(testResult, new ParseTreeValue(result, typeName), $"{typeName}: '{op}' operator Failed");
            }
            else if (op.Equals("Pow"))
            {
                var testResult = ParseTreeValue.Pow(LHS, RHS );
                Assert.AreEqual(testResult, new ParseTreeValue(result, typeName), $"{typeName}: '{op}' operator Failed");
            }
            else if (op.Equals("Mod"))
            {
                var testResult = LHS % RHS;
                Assert.AreEqual(testResult, new ParseTreeValue(result, typeName), $"{typeName}: '{op}' operator Failed");
            }
            else
            {
                Assert.IsFalse(true, $"operation: {op} - has no test code");
            }
        }

        [TestCase("z * b", "5 To 10", ExpectedResult = "Long")]
        [TestCase("z * c", "5 To 10", ExpectedResult = "Long")]
        //TODO: Resolve[TestCase("CStr(z) & d", "5 To 10", ExpectedResult = "String")]
        //TODO: Resolve[TestCase("z & d", "d To dd", ExpectedResult = "String")]
        //TODO: determine if this is a legal statement[TestCase(@"z & ""45""", "5 To 10", ExpectedResult = "Long")]
        [TestCase("Random() > 0.5", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("True", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("z And True", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("z And j > 0.00", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("TestValueLong()", "5 To 10", ExpectedResult = "Long")]
        [TestCase("v", "b To bb", ExpectedResult = "Long")]
        [TestCase("v", "5 To 100", ExpectedResult = "Long")]
        [TestCase("v", "5 To 45.6", ExpectedResult = "Double")]
        [TestCase("v", "ToLong( j ) * bb * Random() + b ^ 4.5", ExpectedResult = "Double")]
        [TestCase("v", @"ToLong(""Whaaaat"") * bb * Random() + b ^ 4.5", ExpectedResult = "Double")]
        [TestCase("hint&", "5.0 To 45.6", ExpectedResult = "Long")]
        [TestCase("Sunday", "5.0 To 45.6", ExpectedResult = "Long")]
        [Category("Inspections")]
        public string UnreachableCaseInspection_DetermineSelectCaseType(string selectExpr, string firstCaseExpr)
        {
            string inputCode =
@"
        Private Enum Weekday
            Sunday = 1
            Monday = 2
            Tuesday = 3
            Wednesday = 4
            Thursday = 5
            Friday = 6
            Saturday = 7
        End Enum
        Private const b As Long = 55
        Private const bb As Long = 100
        Private const c As Double = 0.0023
        Private const cc As Double = 0.509
        Private const d As String = ""Bar""
        Private const dd As String = ""Foo""

        Private Function ToLong(val As Variant) As Long
           ToLong = Clng( val )
        End Function

        Private Function Random() As Double
            Random = VBA.Rnd()
        End Function

        Private Function TestValueLong() As Long
            TestValueLong = 5
        End Function

        Sub Foo(z As Long, j As Double, m As String, v as Variant, w as Boolean)

        Dim hint&
        hint& = 25

        Select Case <selectExpr>
          Case <firstCaseExpr>
            'OK
          Case Else
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<selectExpr>", selectExpr);
            inputCode = inputCode.Replace("<firstCaseExpr>", firstCaseExpr);
            return GetSelectCaseEvaluationType(inputCode);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_StringsInTheMixConvertable()
        {
            const string inputCode =
@"
        Private Function ToLong(val As Variant) As Long
            ToLong = 5
        End Function

        Sub Foo(z As Long, s As String)

        Select Case z + ToLong(s)
            Case ""105""
            'OK
            Case 55
            'Unreachable
            Case 55
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";

            var result = GetSelectCaseEvaluationType(inputCode);
            Assert.AreEqual(Tokens.Long, result);
        }

        [TestCase("50 To 100", 50, 100)]
        [TestCase("fromVal To toVal", 50, 100)]
        [TestCase("50.25 To 100.49", 50, 100)]
        [TestCase("True To False", -1, 0)]
        [TestCase("False To True", 0, -1)]
        [TestCase(@"""50"" To ""100""", 50, 100)]
        [TestCase("100 To 50", 100, 50)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectStmtParseTreeValues(string firstCase, long start, long end)
        {
            string inputCode =
@"
        Private Const fromVal As long = 50
        Private Const toVal As Long = 100

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;

            var startContext = ctxtValues.ValueResolvedContexts.Keys.Where(k => k is VBAParser.SelectStartValueContext);
            var endContext = ctxtValues.ValueResolvedContexts.Keys.Where(k => k is VBAParser.SelectEndValueContext);
            Assert.True(startContext.Any(), "Start context not found in Keys");
            Assert.True(endContext.Any(), "End context not found in Keys");
            Assert.AreEqual(ctxtValues.ValueResolvedContexts[startContext.First()].AsLong(), start);
            Assert.AreEqual(ctxtValues.ValueResolvedContexts[endContext.First()].AsLong(), end);
        }

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
        public void UnreachableCaseInspection_SummaryCoverageIsLTClause(string firstCase, long isLTMax, bool isLTE)
        {
            string inputCode =
@"
                Private Const fromVal As long = 500
                Private Const toVal As Long = 1000

                Sub Foo(z As Long)

                Select Case z
                    Case <firstCase>
                    'OK
                End Select

                End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);//, unresolvedValues);
            }
            Assert.AreEqual(isLTMax, summaryCoverage.IsLT.Value, "IsLT value incorrect");
            if (isLTE)
            {
                Assert.AreEqual(true, summaryCoverage.SingleValues.HasCoverage,/*.Values.Any(),*/ "SingleValue not updated");
                //Assert.IsTrue(summaryCoverage.SingleValues.Values.Contains(isLTMax), $"SingleValue is missing Value: {isLTMax}");
                Assert.IsTrue(summaryCoverage.SingleValues.Covers(isLTMax), $"SingleValue is missing Value: {isLTMax}");
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
        public void UnreachableCaseInspection_SummaryCoverageIsGTClause(string firstCase, long isGTMin, bool isGTE)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            var IsGTMin = summaryCoverage.IsGT.Value;
            Assert.AreEqual(isGTMin, IsGTMin, "IsGT value incorrect");
            if (isGTE)
            {
                Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
                Assert.IsTrue(summaryCoverage.SingleValues.Values.Contains(isGTMin), $"SingleValue is missing Value: {isGTMin}");
            }
        }

        //[TestCase("Is < x", 0)]
        //[TestCase("z < x", 0)]
//        [Category("Inspections")]
//        public void UnreachableCaseInspection_SummaryCoverageIsLTClauseVariable(string firstCase, long isLTMax)
//        {
//            string inputCode =
//@"
//        Private Const fromVal As long = 500
//        Private Const toVal As Long = 1000

//        Sub Foo(z As Long, x As Long)

//        Select Case z
//            Case <firstCase>
//            'OK
//        End Select

//        End Sub";

//            inputCode = inputCode.Replace("<firstCase>", firstCase);

//            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
//            var ctxtValues = parserTreeValues.Item1;
//            var selectCaseContext = parserTreeValues.Item2;

//            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
//            var summaryCoverage = new SummaryCoverage<long>();
//            var tValues = GetTestValuesLong(ctxtValues);
//            foreach (var caseClause in caseClauses)
//            {
//                summaryCoverage.LoadCoverage(caseClause, tValues);
//            }
//            var coverage = summaryCoverage ;

//            Assert.IsFalse(summaryCoverage.IsLT.HasCoverage);
//        }

        [TestCase("Is = 100", 100)]
        [TestCase("Is = 100.49", 100)]
        [TestCase("Is = 100#", 100)]
        [TestCase("Is = True", -1)]
        [TestCase(@"Is = ""100""", 100)]
        [TestCase("Is = toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsEQClause(string firstCase, long isGTMin)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
            Assert.AreEqual(isGTMin, summaryCoverage.SingleValues.Values.First(), "SingleValue has incorrect Value");
        }

        [TestCase("Is <> 100", 100)]
        [TestCase("Is <> 100.49", 100)]
        [TestCase("Is <> 100#", 100)]
        [TestCase("Is <> True", -1)]
        [TestCase(@"Is <> ""100""", 100)]
        [TestCase("Is <> toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsNEQClause(string firstCase, long isNEQ)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            Assert.IsTrue(summaryCoverage.IsGT.HasCoverage);
            var IsGTMin = summaryCoverage.IsGT;
            Assert.AreEqual(isNEQ, IsGTMin.Value);
            Assert.IsTrue(summaryCoverage.IsLT.HasCoverage);
            var IsLTMax = summaryCoverage.IsLT;
            Assert.AreEqual(isNEQ, IsLTMax.Value);
        }

        [TestCase("z < 100", 100, true)]
        [TestCase("z <= 101", 101, true)]
        [TestCase("100 > z", 100, true)]
        [TestCase("101 >= z", 101, true)]
        [TestCase("z > 300", 300, false)]
        [TestCase("300 < z", 300, false)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageRelationalOp(string firstCase, long valToCheck, bool checkLT)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }
            var coverage = summaryCoverage ;

            if (checkLT)
            {
                Assert.IsTrue(summaryCoverage.IsLT.HasCoverage);
                var IsLTMax = summaryCoverage.IsLT;
                Assert.AreEqual(valToCheck, IsLTMax.Value);
            }
            else
            {
                Assert.IsTrue(summaryCoverage.IsGT.HasCoverage);
                var IsGTMin = summaryCoverage.IsGT;
                Assert.AreEqual(valToCheck, IsGTMin.Value);
            }
        }

        [TestCase("50 * 5", 250)]
        [TestCase("8 / 2", 4)]
        [TestCase("toVal / fromVal", 2)]
        [TestCase("toVal + fromVal", 1500)]
        [TestCase("fromVal - toVal", -500)]
        [TestCase("toVal * True + fromVal / 2", -750)]
        [TestCase("2 ^ 3", 8)]
        [TestCase("9 Mod 4", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageBinaryMathOps(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
            Assert.AreEqual(target, summaryCoverage.SingleValues.Values.First(), "SingleValue has incorrect Value");
        }

        [TestCase("fromVal > 5 And toVal > 20", -1)]
        [TestCase("fromVal > 500000 Or toVal > 20000000", 0)]
        [TestCase("True Xor True", 0)]
        [TestCase("Not fromVal", 0)]
        //[TestCase("10 Eqv 8", -3)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageLogicOps(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
            Assert.AreEqual(target, summaryCoverage.SingleValues.Values.First(), "SingleValue has incorrect Value");
        }

        [TestCase("(fromVal - toVal) * 2", -1000)]
        [TestCase("(((((fromVal) - (toVal)) * (2))))", -1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageParentheses(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
            Assert.AreEqual(target, summaryCoverage.SingleValues.Values.First(), "SingleValue has incorrect Value");
        }

        [TestCase("-fromVal", -500)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageUnaryMinus(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            foreach (var caseClause in caseClauses)
            {
                summaryCoverage.LoadCoverage(caseClause, tValues);
            }

            Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
            Assert.AreEqual(target, summaryCoverage.SingleValues.Values.First(), "SingleValue has incorrect Value");
        }
        [TestCase("BitCountMaxValues.max1Bits", 1)]
        [TestCase("BitCountMaxValues.max2Bits", 3)]
        [TestCase("BitCountMaxValues.max3Bits", 7)]
        [TestCase("BitCountMaxValues.max4Bits", 15)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumMemberAccess(string firstCase, long value)
        {
            string inputCode =
@"
        private Enum BitCountMaxValues
            max1Bits = 2 ^ 0
            max2Bits = 2 ^ 1 + max1Bits
            max3Bits = 2 ^ 2 + max2Bits
            max4Bits = 2 ^ 3 + max3Bits
        End Enum

        Sub Foo(z As BitCountMaxValues)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            var caseVals = new List<long>() { value };

            inputCode = inputCode.Replace("<firstCase>", firstCase);

            string result = string.Empty;
            ContextValueVisitor ptValues = null;
            //ContextValueDecorator wrapper = null;
            VBAParser.SelectCaseStmtContext selectCaseContext = null;

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                ptValues = new ContextValueVisitor(state, Tokens.Long);
                selectCaseContext = GetDescendent<VBAParser.SelectCaseStmtContext>(state.ParseTrees.First().Value);
                //wrapper = new ContextValueDecorator(selectCaseContext);
                //wrapper.Accept(ptValues);
                selectCaseContext.Accept(ptValues);
            }

            var summaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ptValues);
            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>();
            var caseClauseWrappers = new List<CaseClauseWrapper<long>>();
            var caseClauseWrapper = new CaseClauseWrapper<long>(caseClauses.First(), tValues);
            summaryCoverage = caseClauseWrapper.SummaryCoverage;

            Assert.AreEqual(true, summaryCoverage.SingleValues.Values.Any(), "SingleValue not updated");
            Assert.IsTrue(summaryCoverage.SingleValues.Values.All(sv => caseVals.Contains(sv)));
        }

        [TestCase("IsLT=45,Range=20:70", "IsLT=45", "Range=20:70")]
        [TestCase("Range=20:70,IsLT=45", "IsLT=45", "Range=20:70")]
        [TestCase("IsLT=45,Range=20:70", "Range=10:70", "IsLT=45")]
        [TestCase("IsLT=45,IsGT=105,Range=20:70", "IsLT=45,Single=200", "IsGT=105,Range=20:70")]
        [TestCase("IsLT=45,IsGT=205,Range=20:70,Single=200", "IsLT=45,IsGT=205,Range=20:70", "Single=200")]
        [TestCase("Range=60:80", "Range=20:70,Range=65:100", "")]
        [TestCase("Single=17", "Range=1:4,Range=7:9,Range=15:20", "")]
        [TestCase("Range=101:149", "Range=150:250,Range=1:100",  "Range=101:149")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_RemovalRangeClauses(string firstCase, string secondCase, string expectedClauses)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase, expectedClauses });
            var sumClauses = new List<SummaryCoverage<long>>();
            foreach(var id in caseToRanges)
            {
                var newSummary = new SummaryCoverage<long>();
                newSummary = CreateTestSummaryCoverageLong(id.Value, newSummary);
                sumClauses.Add(newSummary);
            }

            var candidateClause = sumClauses[0];
            var existingClauses = sumClauses[1];
            var check = sumClauses[2];
            
            var nonDuplicates = candidateClause.RemoveCoverageRedundantTo(existingClauses);
            Assert.AreEqual(check.ToString(), nonDuplicates.ToString());
        }

        [TestCase("IsLT=40,IsGT=40", "Range=35:45", "Long")]
        [TestCase("IsLT=40,IsGT=44", "Range=35:45", "Long")]
        [TestCase("IsLT=40,IsGT=40", "Single=40", "Long")]
        [TestCase("IsGT=240,Range=150:239", "Single=240, Single=0,Single=1,Range=2:150", "Byte")]
        [TestCase("Range=151:255", "Single=150, Single=0,Single=1,Range=2:149", "Byte")]
        [TestCase("IsLT=13,IsGT=30,Range=30:100", "Single=13,Single=14,Single=15,Single=16,Single=17,Single=18,Range=19:30", "Long")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CoversAll(string firstCase, string secondCase, string typeName)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase });
            var sumClauses = new List<SummaryCoverage<long>>();
            var compareMin = CompareExtents.LONGMIN;
            var compareMax = CompareExtents.LONGMAX;
            var cumClause = new SummaryCoverage<long>();
            foreach (var id in caseToRanges)
            {
                var newSummary = new SummaryCoverage<long>();
                newSummary = CreateTestSummaryCoverageLong(id.Value, newSummary);
                if (typeName.Equals(Tokens.Integer))
                {
                    compareMin = CompareExtents.INTEGERMIN;
                    compareMax = CompareExtents.INTEGERMAX;
                    newSummary.ApplyExtents(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX);
                }
                if (typeName.Equals(Tokens.Byte))
                {
                    compareMin = CompareExtents.BYTEMIN;
                    compareMax = CompareExtents.BYTEMAX;
                    newSummary.ApplyExtents(CompareExtents.BYTEMIN, CompareExtents.BYTEMAX);
                }
                var diff = newSummary.RemoveCoverageRedundantTo(cumClause);
                cumClause.Add(diff);
            }
            Assert.IsTrue(cumClause.CoversAllValues);
        }

        [TestCase("IsLT=True, IsGT=True", "Single=False", "")]
        [TestCase("IsLT=False, IsGT=False", "Single=True", "")]
        [TestCase("Single=True, Single=False", "Single=True", "Single=False")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryClausesBoolean(string firstCase, string secondCase, string expectedClauses)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase, expectedClauses });

            var sumClauses = new List<SummaryCoverage<bool>>();
            foreach (var id in caseToRanges)
            {
                var newSummary = CreateTestSummaryCoverageBoolean(id.Value);
                sumClauses.Add(newSummary);
            }

            var candidateClause = sumClauses[0];
            var existingClauses = sumClauses[1];
            var check = sumClauses[2];

            var nonDuplicates = candidateClause.RemoveCoverageRedundantTo(existingClauses);
            Assert.AreEqual(check, nonDuplicates);
        }

        [TestCase("Range=101:149,Range=1:100", "Range=150:250", "Range=1:250")]
        [TestCase("Range=101:149,Range=1:100", "Range=150:250,Range=25:249", "Range=1:250")]
        [TestCase("Range=101:149", "Range=15:148", "Range=15:149")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CombineRangesInteger(string firstCase, string secondCase, string expectedClauses)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase, expectedClauses });
            var sumClauses = new List<SummaryCoverage<long>>();
            foreach (var id in caseToRanges)
            {
                var newSummary = new SummaryCoverage<long>();
                newSummary = CreateTestSummaryCoverageLong(id.Value, newSummary);
                sumClauses.Add(newSummary);
            }

            var firstClause = sumClauses[0];
            var secondClause = sumClauses[1];
            var expected = sumClauses[2];

            firstClause.Add(secondClause);

            Assert.AreEqual(expected.ToString(), firstClause.ToString());
        }

        [TestCase("Range=101.45:149.00007", "Range=101.57:110.63", "Range=101.45:149.00007")]
        [TestCase("Range=101.45:149.0007", "Range=15.67:148.9999", "Range=15.67:149.0007")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CombineRangesRational(string firstCase, string secondCase, string expectedClauses)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase, expectedClauses });
            var sumClauses = new List<SummaryCoverage<double>>();
            foreach (var id in caseToRanges)
            {
                var newSummary = new SummaryCoverage<double>();
                newSummary = CreateTestSummaryCoverageDouble(id.Value, newSummary);
                sumClauses.Add(newSummary);
            }

            var firstClause = sumClauses[0];
            var secondClause = sumClauses[1];
            var expected = sumClauses[2];

            firstClause.Add(secondClause);

            Assert.AreEqual(expected.ToString(), firstClause.ToString());
        }

        //private static string StringizeSummaryClauseXXX<T>(SummaryCoverage<T> summaryClause) where T: IComparable<T>
        //{
        //    var result = string.Empty;
        //    if (summaryClause.IsLT.HasCoverage)
        //    {
        //        result = result + "IsLT=" + summaryClause.IsLT.Value.ToString() + ",";
        //    }
        //    if (summaryClause.IsGT.HasCoverage)
        //    {
        //        result = result + "IsGT=" + summaryClause.IsGT.Value.ToString() + ",";
        //    }
        //    foreach (var range in summaryClause.Ranges.RangeClauses)
        //    {
        //        result = result + "Range=" + range.Start.ToString() + ":" + range.End.ToString() + ",";
        //    }
        //    foreach (var singleVal in summaryClause.SingleValues.Values)
        //    {
        //        result = result + "Single=" + singleVal.ToString() + ",";
        //    }
        //    if(result.Length > 0)
        //    {
        //        return result.Remove(result.Length - 1);
        //    }
        //    return string.Empty;
        //}

        [TestCase("Single=45000", "Single=-50000", "Integer")]
        [TestCase("IsGT=45000", "IsLT=-50000", "Integer")]
        //[TestCase("Range=-450000:-45000", "Range=33000:50000", "Integer")]
        [TestCase("IsGT=45000", "IsLT=-50000", "Byte")]
        //[TestCase("Range=-5:-2", "Range=300:400", "Byte")]
        //[TestCase("Range=250:300", "Range=-10:55", "Byte")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ApplyExtentsPostLoad(string firstCase, string secondCase, string typeName)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase });
            var sumClauses = new List<SummaryCoverage<long>>();
            var compareMin = CompareExtents.LONGMIN;
            var compareMax = CompareExtents.LONGMAX;
            foreach (var id in caseToRanges)
            {
                var newSummary = new SummaryCoverage<long>();
                newSummary = CreateTestSummaryCoverageLong(id.Value, newSummary);
                if (typeName.Equals(Tokens.Integer))
                {
                    compareMin = CompareExtents.INTEGERMIN;
                    compareMax = CompareExtents.INTEGERMAX;
                    newSummary.ApplyExtents(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX);
                }
                if (typeName.Equals(Tokens.Byte))
                {
                    compareMin = CompareExtents.BYTEMIN;
                    compareMax = CompareExtents.BYTEMAX;
                    newSummary.ApplyExtents(CompareExtents.BYTEMIN, CompareExtents.BYTEMAX);
                }
                sumClauses.Add(newSummary);
            }

            foreach (var summaryClause in sumClauses)
            {
                if (summaryClause.IsLT.HasCoverage)
                {
                    Assert.IsTrue(summaryClause.IsLT.Value.CompareTo(compareMin) == 0, "IsLT value is incorrect");
                }
                if (summaryClause.IsGT.HasCoverage)
                {
                    Assert.IsTrue(summaryClause.IsGT.Value.CompareTo(compareMax) == 0, "IsGT value is incorrect");
                }
                if (summaryClause.Ranges.HasCoverage)
                {
                    Assert.IsFalse(summaryClause.Ranges.RangeClauses.Any(rg => rg.Start.CompareTo(compareMin) < 0 || rg.End.CompareTo(compareMax) > 0), "Ranges contain an incorrect value");
                }
            }
        }

        [TestCase("Single=45000", "Single=-50000", "Integer")]
        [TestCase("IsGT=45000", "IsLT=-50000", "Integer")]
        [TestCase("Range=-450000:-45000", "Range=33000:50000", "Integer")]
        [TestCase("IsGT=45000", "IsLT=-50000", "Byte")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ApplyExtentsPreLoad(string firstCase, string secondCase, string typeName)
        {
            var caseToRanges = CasesToRanges(new string[] { firstCase, secondCase });
            var sumClauses = new List<SummaryCoverage<long>>();
            var compareMin = CompareExtents.LONGMIN;
            var compareMax = CompareExtents.LONGMAX;
            foreach (var id in caseToRanges)
            {
                var newSummary = new SummaryCoverage<long>();
                if (typeName.Equals(Tokens.Integer))
                {
                    compareMin = CompareExtents.INTEGERMIN;
                    compareMax = CompareExtents.INTEGERMAX;
                    newSummary.ApplyExtents(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX);
                    newSummary = CreateTestSummaryCoverageLong(id.Value, newSummary);
                }
                if (typeName.Equals(Tokens.Byte))
                {
                    compareMin = CompareExtents.BYTEMIN;
                    compareMax = CompareExtents.BYTEMAX;
                    newSummary.ApplyExtents(CompareExtents.BYTEMIN, CompareExtents.BYTEMAX);
                    newSummary = CreateTestSummaryCoverageLong(id.Value, newSummary);
                }
                sumClauses.Add(newSummary);
            }

            foreach (var summaryClause in sumClauses)
            {
                if (summaryClause.IsLT.HasCoverage)
                {
                    Assert.IsTrue(summaryClause.IsLT.Value.CompareTo(compareMin) == 0, "IsLT value is incorrect");
                }
                if (summaryClause.IsGT.HasCoverage)
                {
                    Assert.IsTrue(summaryClause.IsGT.Value.CompareTo(compareMax) == 0, "IsGT value is incorrect");
                }
            }
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

        private SummaryCoverage<long> CreateTestSummaryCoverageLong(List<string> annotations, SummaryCoverage<long> result)
        {
            //var result = new SummaryCoverage<long>();
            foreach (var item in annotations)
            {
                var element = item.Trim().Split(new string[] { "=" }, StringSplitOptions.None);
                if (element[0].Equals(string.Empty) || element.Count() < 2)
                {
                    continue;
                }
                var clauseType = element[0];
                var clauseExpression = element[1];
                if (clauseType.Equals("IsLT"))
                {
                    result.SetIsLT(long.Parse(clauseExpression));
                }
                else if (clauseType.Equals("IsGT"))
                {
                    result.SetIsGT(long.Parse(clauseExpression));
                }
                else if (clauseType.Equals("Range"))
                {
                    var startEnd = clauseExpression.Split(new string[] { ":" }, StringSplitOptions.None);
                    result.AddRange(long.Parse(startEnd[0]), long.Parse(startEnd[1]));
                }
                else if (clauseType.Equals("Single"))
                {
                    result.Add(long.Parse(clauseExpression));
                }
            }
            return result;
        }

        private SummaryCoverage<double> CreateTestSummaryCoverageDouble(List<string> annotations, SummaryCoverage<double> result)
        {
            foreach (var item in annotations)
            {
                var element = item.Trim().Split(new string[] { "=" }, StringSplitOptions.None);
                if (element[0].Equals(string.Empty) || element.Count() < 2)
                {
                    continue;
                }
                var clauseType = element[0];
                var clauseExpression = element[1];
                if (clauseType.Equals("IsLT"))
                {
                    result.SetIsLT(double.Parse(clauseExpression));
                }
                else if (clauseType.Equals("IsGT"))
                {
                    result.SetIsGT(double.Parse(clauseExpression));
                }
                else if (clauseType.Equals("Range"))
                {
                    var startEnd = clauseExpression.Split(new string[] { ":" }, StringSplitOptions.None);
                    result.AddRange(double.Parse(startEnd[0]), double.Parse(startEnd[1]));
                }
                else if (clauseType.Equals("Single"))
                {
                    result.Add(double.Parse(clauseExpression));
                }
            }
            return result;
        }

        private SummaryCoverage<bool> CreateTestSummaryCoverageBoolean(List<string> annotations)
        {
            var result = new SummaryCoverage<bool>();
            foreach (var item in annotations)
            {
                var element = item.Split(new string[] { "=" }, StringSplitOptions.None);
                if (element[0].Equals(string.Empty) || element.Count() < 2)
                {
                    continue;
                }
                var clauseType = element[0];
                var clauseExpression = element[1];
                if (clauseType.Equals("IsLT"))
                {
                    result.SetIsLT(bool.Parse(clauseExpression));
                }
                else if (clauseType.Equals("IsGT"))
                {
                    result.SetIsGT(bool.Parse(clauseExpression));
                }
                else if (clauseType.Equals("Range"))
                {
                    var startEnd = clauseExpression.Split(new string[] { ":" }, StringSplitOptions.None);
                    result.AddRange(bool.Parse(startEnd[0]), bool.Parse(startEnd[1]));
                }
                else if (clauseType.Equals("Single"))
                {
                    result.Add(bool.Parse(clauseExpression));
                }
            }
            return result;
        }

        [TestCase("toVal_fromVal_500", 1)]
        [TestCase("Is < toVal_fromVal_500", 2)]
        [TestCase("toVal_fromVal To toVal_750", 1)]
        [TestCase("0 To 50_25 To 75_20 To 51", 1)]
        [TestCase("Is > 0_fromVal To toVal_-100", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummarizeResults(string allCases, long expected)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
                'foo
            Case <secondCase>
                'bar
            Case <thirdCase>
                'stuff
            Case Else
                'final stuff
        End Select

        End Sub";
            var separator = new string[] { "_" };
            var firstCase = allCases.Split(separator, StringSplitOptions.None)[0];
            var secondCase = allCases.Split(separator, StringSplitOptions.None)[1];
            var thirdCase = allCases.Split(separator, StringSplitOptions.None)[2];

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            inputCode = inputCode.Replace("<secondCase>", secondCase);
            inputCode = inputCode.Replace("<thirdCase>", thirdCase);

            var parserTreeValues = GetParseTreeValues(inputCode, Tokens.Long);
            var ctxtValues = parserTreeValues.Item1;
            var selectCaseContext = parserTreeValues.Item2;

            var unreachableCases = new List<int>();
            var caseClauses = selectCaseContext.GetDescendents<VBAParser.CaseClauseContext>().ToList();
            var overallSummaryCoverage = new SummaryCoverage<long>();
            var tValues = GetTestValuesLong(ctxtValues);
            for (var idx = 0; idx < caseClauses.Count(); idx++)
            {
                var summaryCoverage = new SummaryCoverage<long>();
                summaryCoverage.LoadCoverage(caseClauses[idx], tValues);
                if(idx > 0)
                {
                    var diff = summaryCoverage.RemoveCoverageRedundantTo(overallSummaryCoverage);
                    if(!diff.HasCoverage)
                    {
                        unreachableCases.Add(idx);
                    }
                    else
                    {
                        overallSummaryCoverage.Add(summaryCoverage);
                    }
                }
                else
                {
                    overallSummaryCoverage.Add(summaryCoverage);
                }
            }

            Assert.AreEqual(expected, unreachableCases.Count());
        }

        private string GetSelectCaseEvaluationType(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            string result = string.Empty;
            IParseTree parseTree;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                parseTree = state.ParseTrees.First().Value;
                var cSelect = GetDescendent<VBAParser.SelectCaseStmtContext>(parseTree);
                var qCtxt = new QualifiedContext<ParserRuleContext>(new QualifiedModuleName(), cSelect);
                var wrapper = new SelectStatementInspectionWrapper(qCtxt, state);
                result = wrapper.EvaluationTypeName;
            }
            return result;
        }

        private Tuple<UnreachableCaseInspection, IParseTree> GetInspectionAndParseTree(string inputCode)
        {
            UnreachableCaseInspection inspection;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            string result = string.Empty;
            IParseTree parseTree;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                inspection = new UnreachableCaseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                parseTree = state.ParseTrees.First().Value;
            }
            var ctxt = GetDescendent<VBAParser.SelectCaseStmtContext>(parseTree);
            var listener = inspection.Listener;
            return new Tuple<UnreachableCaseInspection, IParseTree>(inspection, parseTree);
        }

        private Tuple<ContextValueVisitor, VBAParser.SelectCaseStmtContext> GetParseTreeValues(string inputCode, string evaluationTypeName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            string result = string.Empty;
            ContextValueVisitor ptValues = null;
            VBAParser.SelectCaseStmtContext selectCaseCtxt = null;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                ptValues = new ContextValueVisitor(state);
                selectCaseCtxt = GetDescendent<VBAParser.SelectCaseStmtContext>(state.ParseTrees.First().Value);
                ptValues.EvaluationTypeName = evaluationTypeName;
                selectCaseCtxt.Accept(ptValues);
            }
            return new Tuple<ContextValueVisitor, VBAParser.SelectCaseStmtContext>(ptValues, selectCaseCtxt);
        }

        ContextValueResults<long> GetTestValuesLong(ContextValueVisitor ctxtValues)
        {
            var result = ctxtValues.ResultsAsLong();
            result.Extents.MinMax(CompareExtents.LONGMIN, CompareExtents.LONGMAX);
            return result;
        }

        private T GetDescendent<T>(IParseTree pt)
        {
            for(var idx = 0; idx < pt.ChildCount; idx++)
            {
                var child = pt.GetChild(idx);
                if(child is T)
                {
                    return (T)child;
                }
                if(child.ChildCount > 0)
                {
                    var result =  GetDescendent<T>(child);
                    if(result is T)
                    {
                        return result;
                    }
                }
            }
            return default;
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
        //Negative values
        [TestCase("Long", "-450000", "850000")]
        [TestCase("Integer", "-4500", "8500")]
        [TestCase("Double", "-45000.345", "55000.25")]
        [TestCase("Single", "-45.345", "55.25")]
        [TestCase("Currency", "-4.34578", "5.25869")]
        [TestCase("Boolean", "-55", "0")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SingleUnreachableAllTypes(string type, string value1, string value2)
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
        public void UnreachableCaseInspection_ExceedsLimits(string type, string value1, string value2)
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
        [TestCase("x Eqv 1")]
        [TestCase("Not x")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LogicalOpSelectCase(string booleanOp)
        {
            string inputCode =
@"Sub Foo(x As Long)
        Select Case <boolOp>
            Case True
            'OK
            Case False 
            'OK
            Case -5
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<boolOp>", booleanOp);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestCase("Is > 8", "12", "9")]
        [TestCase("Is >= 8", "12", "8")]
        [TestCase("Is < 8", "-56", "7")]
        [TestCase("Is <= 8", "-56", "8")]
        [TestCase("Is <> 8", "-56", "5000")]
        [TestCase("Is = 8", "16 / 2", "4 * 2")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_IsStmt(string isStmt, string unreachableValue1, string unreachableValue2)
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

        [TestCase("Dim Hint$\r\nSelect Case Hint$", @"""Here"" To ""Eternity""", @"""Forever""")] //String
        [TestCase("Dim Hint#\r\nHint#= 1.0\r\nSelect Case Hint#", "10.00 To 30.00", "20.00")] //Double
        [TestCase("Dim Hint!\r\nHint! = 1.0\r\nSelect Case Hint!", "10.00 To 30.00", "20.00")] //Single
        [TestCase("Dim Hint%\r\nHint% = 1\r\nSelect Case Hint%", "10 To 30", "20")] //Integer
        [TestCase("Dim Hint&\r\nHint& = 1\r\nSelect Case Hint&", "1000 To 3000", "2000")] //Long
        [Category("Inspections")]
        public void UnreachableCaseInspection_TypeHint(string typeHintExpr, string firstCase, string secondCase)
        {
            string inputCode =
@"
        Sub Foo()

        <typeHintExprAndSelectCase>
            Case <firstCaseVal>
            'OK
            Case <secondCaseVal>
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<typeHintExprAndSelectCase>", typeHintExpr);
            inputCode = inputCode.Replace("<firstCaseVal>", firstCase);
            inputCode = inputCode.Replace("<secondCaseVal>", secondCase);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestCase("Long", "Is < 5", "x > -5000")]
        [TestCase("Long", "Is <> 4", "4")]
        [TestCase("Long", "Is <> -4", "4 - 8")]
        [TestCase("Long", "x > -5000", "Is < 1")]
        [TestCase("Long", "-5000 < x", "Is < 1")]
        [TestCase("Integer", "x <> 40", "35 To 45")]
        [TestCase("Double", "x > -5000.0", "Is < 1.7")]
        [TestCase("Single", "x > -5000.0", "Is < 1.7")]
        [TestCase("Currency", "x > -5000.0", "Is < 1.7")]
        [TestCase("Boolean", "-5000", "False")]
        [TestCase("Boolean", "True", "0")]
        [TestCase("Boolean", "50", "0")]
        //[TestCase("Boolean", "Is > -1", "-10")]
        //[TestCase("Boolean", "Is < -100", "Is > -10")]
        //[TestCase("Boolean", "Is < 0", "0")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CoversAllVariousTypes(string type, string firstCase, string secondCase)
        {
            string inputCode =
@"Sub Foo(x As <Type>)

        Select Case x
            Case <firstCase>
            'OK
            Case <secondCase>
            'OK
            Case 45 * 12
            'Unreachable
            Case 500 To 700
            'Unreachable
            Case Else
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<Type>", type);
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            inputCode = inputCode.Replace("<secondCase>", secondCase);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2, caseElse: 1);
        }

        [TestCase("0 To 10")]
        //[TestCase("Is < 1")]
        //[TestCase("-10 To 5")] -> is "True To True"
        //[TestCase("5 To -10")] -> is "True To True"
        [TestCase("True To False")]
        [TestCase("False To True")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanSingleStmtCoversAll(string firstCase)
        {
            string inputCode =
@"Sub Foo(x As Boolean)

        Select Case x
            Case <firstCase>
            'OK
            Case False
            'unreachable
            Case Else
            'Unreachable
        End Select

        End Sub";
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        //TODO: These tests should always fail until at text only capability is added
//        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 2 = 49, x ^ 3 = 8")]
//        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30")]
//        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "(CLng(VBA.Rnd() * 100) * x) < 30, x ^ 2 = 49")]
//        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 3 = 8")]
//        [Category("Inspections")]
//        public void UnreachableCaseInspection_NoInspectionTextCompareOnly(string complexClause1, string complexClause2)
//        {
//            string inputCode =
//@"Sub Foo(x As Long)

//        Select Case x
//            Case <complexClause1>
//            'OK
//            Case <complexClause2>
//            'Unreachable - detected by text compare of range clause(s)
//        End Select

//        End Sub";
//            inputCode = inputCode.Replace("<complexClause1>", complexClause1);
//            inputCode = inputCode.Replace("<complexClause2>", complexClause2);
//            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
//        }

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
        public void UnreachableCaseInspection_CaseClauseHasBinaryMathOp(string type, string mathOp, string unreachable)
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

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_PowOpEvaluationAlgebraNoDetection()
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

        //TODO: this has a mathop in the SelectStmt, but there is no Is clauses...evaluate?
        [Ignore("Infer evaluation type")]
        //[Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NumberRangeConstants()
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

        [TestCase(@"1 To ""Forever""")]
        [TestCase(@"""Fifty-Five"" To 1000")]
        [TestCase("z To 1000")]
        [TestCase("50 To z")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NumberRangeMixedTypes(string firstCase)
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
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, mismatch: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NumberRangeCummulativeCoverage()
        {
            const string inputCode =
@"Sub Foo(x as Long)

        Select Case x
            Case 150 To 250
            'OK
            Case 1 To 100
            'OK
            Case 101 To 149
            'OK
            Case 25 To 249 
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NumberRangeHighToLow()
        {
            const string inputCode =
@"Sub Foo(x as Long)

        Select Case x
            Case 100 To 1
            'OK
            Case 50
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseElseIsClausePlusRange()
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
        public void UnreachableCaseInspection_CaseElseIsClausePlusRangeAndSingles()
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
        public void UnreachableCaseInspection_NestedSelectCase()
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
        public void UnreachableCaseInspection_NestedSelectCases()
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
        public void UnreachableCaseInspection_NestedSelectCaseSUnreachable()
        {
            const string inputCode =
@"Sub Foo(x As String, z As String)

        'Const x As String = ""Foo""
        'Const z As String = ""Bar""

        Select Case x
            Case ""Foo"", ""Bar""
            'OK
            Case ""Foo""
            'Unreachable
            Case ""Food""
            Select Case  z
                Case ""Foo"", ""Bar"",""Goo""
                'OK
                Case ""Bar""
                'Unreachable
                Case ""Foo""
                'Unreachable
                Case ""Goo""
                'Unreachable
            End Select
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 4);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SimpleLongCollisionConstantEvaluation()
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
        //TODO: Can this test reasonaby return Double as the evaluation type?
        [Ignore("Infer evaluation type")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_MixedSelectCaseTypes()
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
        public void UnreachableCaseInspection_ExceedsIntegerButIncludesAccessibleValues()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

        Select Case x
            Case 10,11,12
            'OK
            Case 15, 40000
            'Exceeds Integer value - but other value makes case reachable....no Error
            Case x < 4
            'OK
            Case -50000
            'Exceeds Integer values
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_IntegerWithDoubleValue()
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
        public void UnreachableCaseInspection_VariantSelectCase()
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

        //TODO: can we infer Double here...there are no Is statements
        [Ignore("Infer inspection type")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_VariantSelectCaseInferFromConstant()
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
        public void UnreachableCaseInspection_VariantSelectCaseInferFromConstant2()
        {
            const string inputCode =
@"Sub Foo( x As Variant)

        private Const TheValue As Double = 45.678
        private Const TheUnreachableValue As Long = 77

        Select Case x
            Case x > TheValue
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
        public void UnreachableCaseInspection_BuiltInSelectCase()
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

        [TestCase("True", "x <> False", 2)]
        [TestCase("x >= True", "False", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanRelationalOps(string firstCase, string secondCase, int expected)
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

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LongCollisionIndeterminateCase()
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

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LongCollisionMultipleVariablesSameType()
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
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LongCollisionMultipleVariablesDifferentType()
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
            'OK - this gives a false positive when evaluated as if 'x' or 'y' is the only select case variable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Ignore("One of the Twelve")]
        //[Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LongCollisionVariableAndConstantDifferentType()
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
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [Ignore("One of the Twelve")]
        //[Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LongCollisionUnaryMathOperation()
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
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElseInvertBooleanRange()
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
        public void UnreachableCaseInspection_StringWhereLongShouldBe()
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
        public void UnreachableCaseInspection_MixedTypes()
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
        public void UnreachableCaseInspection_StringWhereLongShouldBeIncludeLongAsString()
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

//        [Test] -> Replicated
//        [Category("Inspections")]
//        public void UnreachableCaseInspection_MultipleRanges()
//        {
//            const string inputCode =
//@"Sub Foo(x As Long)

//        Select Case x
//            Case 1 To 4, 7 To 9, 11, 13, 15 To 20
//            'OK
//            Case 8
//            'Unreachable
//            Case 11
//            'Unreachable
//            Case 17
//            'Unreachable
//            Case 21
//            'Reachable
//        End Select

//        End Sub";
//            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
//        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CascadingIsStatements()
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
        public void UnreachableCaseInspection_CascadingIsStatementsGT()
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
        public void UnreachableCaseInspection_IsStatementUnreachableGT()
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
        public void UnreachableCaseInspection_IsStatementUnreachableLT()
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
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachableUsingIs()
        {
            const string inputCode =
@"Sub Foo(z As Long)

        Select Case z
            Case z <> 5 
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
        public void UnreachableCaseInspection_CaseClauseHasParens()
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
        public void UnreachableCaseInspection_CaseClauseHasMultipleParens()
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

        [Ignore("One of the Twelve")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectCaseHasMultOpWithFunction()
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
        public void UnreachableCaseInspection_CaseClauseHasMultOpInParens()
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
        public void UnreachableCaseInspection_CaseClauseHasMultOp2Constants()
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
        public void UnreachableCaseInspection_EnumerationNumberRangeNoDetection()
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
            'Unreachable
            Case Weekday.Wednesday
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumerationNumberRangeNonConstant()
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
        public void UnreachableCaseInspection_EnumerationLongCollision()
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
        public void UnreachableCaseInspection_EnumerationNumberRangeConflicts()
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
        public void UnreachableCaseInspection_EnumerationNumberCaseElse()
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
                    Case z <> Apple
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
        public void UnreachableCaseInspection_CaseElseByte()
        {
            const string inputCode =
@"
        Sub Foo(z As Byte)

        Select Case z
            Case z >= 2
            'OK
            Case 0,1
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseElseByteMultipleCases()
        //        {
        //            const string inputCode =
        //@"
        //        Sub Foo(z As Byte)

        //        Select Case z
        //            Case z >= 240
        //            'OK
        //            Case 0,1
        //            'OK
        //            Case Is < 100
        //            'OK
        //            Case 150 To 240
        //            'OK
        //            Case 100 To 228
        //            'OK
        //            Case Else
        //            'Unreachable
        //        End Select

        //        End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_RangeCollisionsAggregateClauses()
        //        {
        //            const string inputCode =
        //@"
        //        Sub Foo(z As Long)

        //        Select Case z
        //            Case z > 30
        //            'OK
        //            Case 14,15,16,17,18,19 To 30
        //            'OK
        //            Case 30 To 100
        //            'Unreachable
        //            Case Is <= 13
        //            'OK   
        //            Case Else
        //            'Unreachable
        //        End Select

        //        End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        //        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstantReferenceExpr()
        {
            const string inputCode =
@"
        private Const maxValue As Long = 5000

        Sub Foo(z As Long)

        Select Case ( z * 3 ) - 2
            Case z > maxValue
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
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstantIndeterminantExpression()
        {
            const string inputCode =
@"
        private Const maxValue As Long = 5000

        Sub Foo(z As Long)

        Select Case z
            Case z > maxValue / 2
            'OK
            Case z > maxValue
            'Unreachable
            Case 15
            'OK
            Case 8500
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectCaseIsFunction()
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
        public void UnreachableCaseInspection_SelectCaseIsFunctionWithParams()
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
        public void UnreachableCaseInspection_IsStmtAndNegativeRange()
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
        public void UnreachableCaseInspection_IsStmtAndNegativeRangeWithConstants()
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

            Assert.AreEqual(expected[InspectionsUI.UnreachableCaseInspection_Unreachable], actualUnreachable.Count(), "Unreachable result");
            Assert.AreEqual(expected[InspectionsUI.UnreachableCaseInspection_TypeMismatch], actualMismatches.Count(), "Mismatch result");
            Assert.AreEqual(expected[InspectionsUI.UnreachableCaseInspection_CaseElse], actualUnreachableCaseElses.Count(), "CaseElse result");
        }
    }
}
