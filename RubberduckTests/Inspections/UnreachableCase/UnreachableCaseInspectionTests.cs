using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete.UnreachableCaseInspection;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections.UnreachableCase
{
    [TestFixture]
    public class UnreachableCaseInspectionTests
    {
        private IUnreachableCaseInspectionFactoryProvider _factoryProvider;

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

        private IUnreachableCaseInspectorFactory IUnreachableCaseInspectorFactory => FactoryProvider.CreateIUnreachableInspectorFactory();
        private IParseTreeValueFactory ValueFactory => FactoryProvider.CreateIParseTreeValueFactory();

        [TestCase("Dim Hint$\r\nSelect Case Hint$", @"""Here"" To ""Eternity"",""Forever""", "String")] //String
        [TestCase("Dim Hint#\r\nHint#= 1.0\r\nSelect Case Hint#", "10.00 To 30.00, 20.00", "Double")] //Double
        [TestCase("Dim Hint!\r\nHint! = 1.0\r\nSelect Case Hint!", "10.00 To 30.00,20.00", "Single")] //Single
        [TestCase("Dim Hint%\r\nHint% = 1\r\nSelect Case Hint%", "10 To 30,20", "Integer")] //Integer
        [TestCase("Dim Hint&\r\nHint& = 1\r\nSelect Case Hint&", "1000 To 3000,2000", "Long")] //Long
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectExprTypeHint(string typeHintExpr, string firstCase, string expected)
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

        [TestCase("Dim Hint$", "Hint$", "String")] //String
        [TestCase("Dim Hint#", "Hint#", "Double")] //Double
        [TestCase("Dim Hint!", "Hint!", "Single")] //Single
        [TestCase("Dim Hint%", "Hint%", "Integer")] //Integer
        [TestCase("Dim Hint&", "Hint&", "Long")] //Long
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeHint(string typeHintExpr, string firstCase, string expected)
        {
            string inputCode =
@"
        Sub Foo(x As Variant)

        <typeHintExpr>

        Select Case x
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

        [TestCase("Not x", "x As Long", "Long")]
        [TestCase("x", "x As Long", "Long")]
        [TestCase("x < 5", "x As Long", "Boolean")]
        [TestCase("ToLong(True) * .0035", "x As Byte", "Double")]
        [TestCase("True", "x As Byte", "Boolean")]
        [TestCase("ToString(45)", "x As Byte", "String")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectExpressionType(string selectExpr, string argList, string expected)
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
        [TestCase("ToString(45)", @"""Bar""", "String")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseType(string rangeExpr1, string rangeExpr2, string expected)
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
            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("45", "55", "Integer")]
        [TestCase("45.6", "55", "Double")]
        [TestCase(@"""Test""", @"""55""", "String")]
        [TestCase("True", "y < 6", "Boolean")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeUnrecognizedSelectExpressionType(string rangeExpr1, string rangeExpr2, string expected)
        {
            string inputCode =
$@"
        Sub Foo(x As Collection)
            Dim y As Long
            y = 8
            Select Case x(1)
                Case {rangeExpr1}
                'OK
                Case {rangeExpr2}
                'OK
                Case Else
                'OK
            End Select

        End Sub";

            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("x.Item(2)", "55", "Integer")]
        [TestCase("x.Item(2)", "45.6", "Double")]
        [TestCase("x.Item(2)", @"""Test""", "String")]
        [TestCase("x.Item(2)", "False", "Boolean")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeUnrecognizedCaseExpressionType(string rangeExpr1, string rangeExpr2, string expected)
        {
            string inputCode =
@"
        Sub Foo(x As Collection)
            Select Case x(3)
                Case <rangeExpr1>
                'OK
                Case <rangeExpr2>
                'OK
                Case <rangeExpr2>
                'OK
                Case Else
                'OK
            End Select

        End Sub";

            inputCode = inputCode.Replace("<rangeExpr1>", rangeExpr1);
            inputCode = inputCode.Replace("<rangeExpr2>", rangeExpr2);
            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("x.Item(2)", "True", false)]
        [TestCase("x.Item(2)", "True, False", true)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeUnrecognizedCaseExpressionType2(string rangeExpr1, string rangeExpr2, bool triggersCaseElse)
        {
            string inputCode =
@"
        Sub Foo(x As Collection)
            Select Case x(3)
                Case <rangeExpr1>
                'OK
                Case <rangeExpr2>
                'OK
                Case <rangeExpr2>
                'unreachable
                Case Else
                'Depends on flag
            End Select
        End Sub";

            inputCode = inputCode.Replace("<rangeExpr1>", rangeExpr1);
            inputCode = inputCode.Replace("<rangeExpr2>", rangeExpr2);
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: triggersCaseElse ? 1 : 0);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("95", "55", "Integer")]
        [TestCase("23.2", "45.6", "Double")]
        [TestCase(@"""Foo""", @"""Test""", "String")]
        [TestCase("x < 6", "x > 9", "Boolean")]
        [TestCase("95.7", "55", "Double")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeVariantSelectExpression(string rangeExpr1, string rangeExpr2, string expected)
        {
            string inputCode =
@"
        Sub Foo(x As Variant)
            Select Case x
                Case <rangeExpr1>
                'OK
                Case <rangeExpr2>
                'OK
                Case <rangeExpr2>
                'OK
                Case Else
                'OK
            End Select

        End Sub";

            inputCode = inputCode.Replace("<rangeExpr1>", rangeExpr1);
            inputCode = inputCode.Replace("<rangeExpr2>", rangeExpr2);
            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("Long", "2147486648#", "-2147486649#")]
        [TestCase("Integer", "40000", "-50000")]
        [TestCase("Byte", "256", "-1")]
        [TestCase("Currency", "922337203685490.5808@", "-922337203685477.5809@")]
        [TestCase("Single", "3402824E38", "-3402824E38")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ExceedsLimits(string type, string value1, string value2)
        {
            string inputCode =
$@"Sub Foo(x As {type})

        Const firstVal As {type} = {value1}
        Const secondVal As {type} = {value2}

        Select Case x
            Case firstVal
            'Unreachable
            Case secondVal
            'Unreachable
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_RelationalOpSelectCase()
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        //TODO: we need a test for the other operators <=, > , >= Note: probably belongs 
        //in rangefilter tests
//        [Test]
//        [Category("Inspections")]
//        public void UnreachableCaseInspection_RelationalOp1()
//        {
//            string inputCode =
//@"Sub Foo(x As Long)
//    Select Case x
//        Case x < 100
//        'OK
//        Case 100 > x
//        'Unreachable
//        Case x < 50
//        'Unreachable
//    End Select
//End Sub";
//            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
//            Assert.AreEqual(expectedMsg, actualMsg);
//        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_RelationalOpExpression()
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("fromVal")]
        [TestCase("Not toVal")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LogicalOpUnary(string caseClause)
        {
            string inputCode =
@"Sub Foo(x As Boolean)

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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 2 = 49, x ^ 3 = 8")]
        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30")]
        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "(CLng(VBA.Rnd() * 100) * x) < 30, x ^ 2 = 49")]
        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 3 = 8")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CopyPaste(string complexClause1, string complexClause2)
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase(@"1 To ""Forever""", 1, 1)]
        [TestCase(@"""Fifty-Five"" To 1000", 1, 1)]
        [TestCase("z To 1000", 1, 0)]
        [TestCase("50 To z", 1, 0)]
        [TestCase(@"z To 1000, 95, ""TEST""", 1, 0)]
        [TestCase(@"1 To ""Forever"", 55000", 2, 0)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NumberRangeMixedTypes(string firstCase, int unreachableCount, int mismatchCount)
        {
            string inputCode =
@"Sub Foo(x As Integer, z As String)

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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachableCount, mismatch: mismatchCount);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 4);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NestedSelectCaseStrings()
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NestedSelectCasesUnreachable()
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ExceedsIntegerButIncludesAccessibleValues()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

        Select Case x
            Case -50000
            'Exceeds Integer values and unreachable
            Case 10,11,12
            'OK
            Case 15, 40000
            'Exceeds Integer value - but other value makes case reachable....OK
            Case Is < 4
            'OK
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            Case Is > TheValue
            'OK
            Case 0 To TheValue - 20 '(25.678)
            'OK
            Case TheUnreachableValue
            Unreachable
            Case 55
            'Unreachable
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("True", "Is <= False", 2)]
        [TestCase("Is >= True", "False", 1)]
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: expected);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_InspectButNoResult()
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
            'OK
            Case Else
            'OK
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_DuplicateVariableRange()
        {
            const string inputCode =
@"Sub Foo( x As Long)

        private const y As Double = 0.5

        Select Case x * y
            Case x To 55
            'OK
            Case y > 3000
            'OK
            Case x To 55
            'Unreachable
            Case 95
            'OK
            Case Else
            'OK
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SingleValueRange()
        {
            const string inputCode =
@"Sub Foo( x As Long)

        Select Case x
            Case 55
            'OK
            Case 55 To 55
            'Unreachable
            Case 95
            'OK
            Case Else
            'OK
        End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LongCollisionUnaryMathOperation()
        {
            const string inputCode =
@"
Sub Foo( x As Long, y As Double)
    Select Case -x
        Case x > -3000
        'OK
        Case y > -3000
        'OK
        Case x < y
        'OK - indeterminant
        Case 95
        'OK
        Case x > -3000
        'Ureachable
        Case Else
        'OK
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        //TODO: should there be another error type - Range is high to low?
        [TestCase("False To True", 1, 0)] //<firstCase> is unreachable - malformed
        [TestCase("True To False", 1, 1)] //<firstCase> filters all possible values
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElseInvertBooleanRange(string firstCase, int unreachableCount, int caseElseCount)
        {
            string inputCode =
@"
Private Function Random() As Double
    Random = VBA.Rnd()
End Function

Sub Foo(x As Boolean)
    Select Case Random() > 0.5
        Case <firstCase> 
        'Reachable depends on firstCase
        Case True
        'Reachable depends on firstCase
        Case Else
        'Reachable depends on firstCase
    End Select
End Sub";
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachableCount, caseElse: caseElseCount);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachableUsingIs()
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            'OK
            Case Weekday.Wednesday
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 0, caseElse: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
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
                    Case Is <> Apple
                    'OK
                    Case Apple 
                    'OK     
                    Case Else
                    'unreachable - Guard clause will always be skipped
                    Err.Raise 5, ""MyFunction"", ""Invalid value given for the enumeration.""
                End Select

                End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseElseByte()
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("( z * 3 ) - 2", "Is > maxValue", 2)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstantReferenceExpr(string selectExpr, string firstCase, int expected)
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: expected);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("<")]
        [TestCase(">")]
        [TestCase("<=")]
        [TestCase(">=")]
        [TestCase("<>")]
        [TestCase("=")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_IsStmtVariables(string opSymbol)
        {
            string inputCode =
@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case 95 
            'OK
            Case Is <opSymbol> x
            'OK
            Case -3 To 10 
            'OK - covers True and False
            Case Is <opSymbol> y 
            'OK
            Case Is <opSymbol> x 
            'Unreachable
            Case z < x
            'Unreachable
        End Select

        End Sub";

            inputCode = inputCode.Replace("<opSymbol>", opSymbol);
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("-1 Eqv -1", "True")]
        [TestCase("-1 Imp -1", "True")]
        [TestCase("0 Eqv -1", "False")]
        [TestCase("0 Imp -1", "True")] //=> differs from Eqv
        [TestCase("-1 Eqv 0", "False")]
        [TestCase("-1 Imp 0", "False")]
        [TestCase("0 Eqv 0", "True")]
        [TestCase("0 Imp 0", "True")]
        [TestCase("3 Eqv 0", "-4")]
        [TestCase("3 Imp 0", "-4")]
        [TestCase("0 Eqv 16", "-17")]
        [TestCase("0 Imp 16", "-1")]
        [TestCase("3 Eqv 5", "-7")]
        [TestCase("3 Imp 5", "-3")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ImpEqvOperators(string secondCase, string thirdCase)
        {
            string inputCode =
@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case Is < x 
            'OK
            Case <secondCase>
            'OK
            Case <thirdCase>
            'Unreachable
        End Select

        End Sub";

            inputCode = inputCode.Replace("<secondCase>", secondCase);
            inputCode = inputCode.Replace("<thirdCase>", thirdCase);
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("Eqv")]
        [TestCase("Imp")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ImpEqvOperatorsVariable(string op)
        {
            string inputCode =
@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case Is < x 
            'OK
            Case x <op> y
            'OK
            Case -3 To 10 
            'OK
            Case x <op> y
            'Unreachable
            Case x <op> z
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<op>", op);
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_IntegerDivision()
        {
            string inputCode =
@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case x 
            'OK
            Case 3
            'OK
            Case 10 \ 3
            'Unreachable
        End Select

        End Sub";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        //Issue# 3885
        //this test only proves that the Select Statement is not inspected
        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_BuiltInMember()
        {
            string inputCode =
@"
Option Explicit

Sub FooCount(x As Long)

    Select Case err.Number
        Case ""5903""
            'OK
        Case 5900 + 3
            'Unreachable - but undetected by unit tests, 
        Case 5
            'Unreachable - but undetected by unit tests, 
        Case 4 + 1
            'Unreachable - but undetected by unit tests, 
    End Select

    Select Case x
        Case ""5""
            MsgBox ""Foo""
        Case 2 + 3
            'Unreachable - just to make sure the test finds something 
            MsgBox ""Bar""
    End Select
End Sub
";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_BuiltInMemberInCaseClause()
        {
            string inputCode =
@"
Option Explicit

Sub FooCount(x As Long)

    Select Case x
        Case 5900 + 3
            'OK
        Case err.Number
            'OK - not evaluated
        Case 5903
            'Unreachable
        Case 5900 + 2 + 1
            'Unreachable
    End Select
End Sub
";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        //Issue# 3885 - replicates with UDT rather than a built-in
        [TestCase("Long")]
        [TestCase("Variant")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_MemberAccessor(string propertyType)
        {
            string inputCode =
@"
Option Explicit

Sub AddVariable(testClass As Class1)
    Select Case testClass.AValue
        Case 5903
            'OK
        Case 5900 + 3
            'unreachable
        Case Else
            Exit Sub
    End Select
End Sub";

            string inputClassCode =
@"
Option Explicit

Private myVal As <propertyType>

Public Property Set AValue(val As <propertyType>)
    myVal = val
End Property

Public Property Get AValue() As <propertyType>
    AValue = myVal
End Property
";
            inputClassCode = inputClassCode.Replace("<propertyType>", propertyType);
            var components = new List<(string moduleName, string inputCode)>()
            {
                ("TestModule1", inputCode),
                ("Class1", inputClassCode)
            };

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(components, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("Long")]
        [TestCase("Variant")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_MemberAccessorInCaseClause(string propertyType)
        {
            string inputCode =
@"
Option Explicit

Sub AddVariable(x As Long)
    Select Case x
        Case 300
            'OK
        Case testClass.AValue
            'OK - variable, not value
        Case 150 + 150
            'OK
        Case 3 * 100
            'OK
    End Select
End Sub";

            string inputClassCode =
@"
Option Explicit

Private myVal As <propertyType>

Public Property Set AValue(val As <propertyType>)
    myVal = val
End Property

Public Property Get AValue() As <propertyType>
    AValue = myVal
End Property
";
            inputClassCode = inputClassCode.Replace("<propertyType>", propertyType);
            var components = new List<(string moduleName, string inputCode)>()
            {
                ("TestModule1",inputCode),
                ("Class1", inputClassCode)
            };

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(components, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("Long = 300")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ConstantInOtherModule(string propertyType)
        {
            string inputCode =
@"
Option Explicit

Sub AddVariable(x As Variant)
    Select Case x
        Case TestModule2.My_CONSTANT
            'OK
        Case 300
            'unreachable
        Case Else
            Exit Sub
    End Select
End Sub";

            string inputModule2Code =
@"
Option Explicit

Public Const MY_CONSTANT As <propertyTypeAndAssignment> 
";
            inputModule2Code = inputModule2Code.Replace("<propertyTypeAndAssignment>", propertyType);
            var components = new List<(string moduleName, string inputCode)>()
            {
                ("TestModule1",inputCode),
                ("TestModule2", inputModule2Code)
            };

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(components, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_DuplicateSelectExpressionVariableInModule()
        {
            string inputCode =
@"
Sub FirstSub(x As Long)
    Select Case x
        Case 55
            MsgBox CStr(x)
        Case 56
            MsgBox CStr(x)
        Case 55
            MsgBox ""Unreachable""
    End Select
End Sub

Sub SecondSub(x As Boolean)
    Select Case x
        Case 55
            MsgBox CStr(x)
        Case 0
            MsgBox CStr(x)
        Case Else
            MsgBox ""Unreachable""
    End Select
End Sub
";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase(@"x Like ", @"""*Bar""", 1)]
        [TestCase(@"y Like ", @"""*Bar""", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_Likes(string caseClauseExpression, string likePattern, int expectedUnreachable)
        {
            string inputCode =
@"
Private Const TEST_CONST As String = ""FooBar""

Sub FirstSub(x As String, y As String)
    Select Case True
        Case <caseClauseExpression><likePattern>
            'OK
        Case TEST_CONST
            'Mismatch
        Case ""CandyBar"" Like <likePattern>
            'OK
        Case ""HandleBar"" Like <likePattern>
            'Unreachable
    End Select
End Sub
";
            inputCode = inputCode.Replace("<caseClauseExpression>", caseClauseExpression);
            inputCode = inputCode.Replace("<likePattern>", likePattern);
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: expectedUnreachable, mismatch: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LikeFiltersToTrue()
        {
            string inputCode =
@"
Private Const TEST_CONST As String = ""FooBar""

Sub FirstSub(x As String)
    Select Case True
        Case x Like ""*""
            'OK
        Case 5 < 6
            'Unreachable
        Case 9 > 8
            'Unreachable
    End Select
End Sub
";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        private static (string expectedMsg, string actualMsg) CheckActualResultsEqualsExpected(string inputCode, int unreachable = 0, int mismatch = 0, int caseElse = 0)
        {
            var components = new List<(string moduleName, string inputCode)>()
            {
                ("TestModule1", inputCode)
            };

            return CheckActualResultsEqualsExpected(components, unreachable, mismatch, caseElse);
        }

        private static (string expectedMsg, string actualMsg) CheckActualResultsEqualsExpected(List<(string moduleName, string inputBlock)> inputCode, int unreachable = 0, int mismatch = 0, int caseElse = 0)
        {
            var expected = new Dictionary<string, int>
            {
                { InspectionResults.UnreachableCaseInspection_Unreachable, unreachable },
                { InspectionResults.UnreachableCaseInspection_TypeMismatch, mismatch },
                { InspectionResults.UnreachableCaseInspection_CaseElse, caseElse },
            };

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected);
            inputCode.ForEach(input => project.AddComponent(input.moduleName, NameToComponentType(input.moduleName), input.inputBlock));
            builder = builder.AddProject(project.Build());
            var vbe = builder.Build();

            IEnumerable<Rubberduck.Parsing.Inspections.Abstract.IInspectionResult> actualResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnreachableCaseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }
            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(InspectionResults.UnreachableCaseInspection_Unreachable));
            var actualMismatches = actualResults.Where(ar => ar.Description.Equals(InspectionResults.UnreachableCaseInspection_TypeMismatch));
            var actualUnreachableCaseElses = actualResults.Where(ar => ar.Description.Equals(InspectionResults.UnreachableCaseInspection_CaseElse));

            var actualMsg = BuildResultString(actualUnreachable.Count(), actualMismatches.Count(), actualUnreachableCaseElses.Count());
            var expectedMsg = BuildResultString(expected[InspectionResults.UnreachableCaseInspection_Unreachable], expected[InspectionResults.UnreachableCaseInspection_TypeMismatch], expected[InspectionResults.UnreachableCaseInspection_CaseElse]);

            return (expectedMsg, actualMsg);
        }

        private static ComponentType NameToComponentType(string name)
        {
            if (name.StartsWith("Class"))
            {
                return ComponentType.ClassModule;
            }
            return ComponentType.StandardModule;
        }

        private static string BuildResultString(int unreachableCount, int mismatchCount, int caseElseCount)
        {
            return  $"Unreachable={unreachableCount}, Mismatch={mismatchCount}, CaseElse={caseElseCount}";
        }

        private string GetSelectExpressionType(string inputCode)
        {
            var selectStmtValueResults = GetParseTreeValueResults(inputCode, out VBAParser.SelectCaseStmtContext selectStmtContext);

            var inspector = IUnreachableCaseInspectorFactory.Create(selectStmtContext, selectStmtValueResults, ValueFactory);
            return ((IUnreachableCaseInspectorTest)inspector).SelectExpressionTypeName;
        }

        private IParseTreeVisitorResults GetParseTreeValueResults(string inputCode, out VBAParser.SelectCaseStmtContext selectStmt)
        {
            selectStmt = null;
            IParseTreeVisitorResults valueResults = null;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var firstParserRuleContext = (ParserRuleContext)state.ParseTrees.Where(pt => pt.Value is ParserRuleContext).First().Value;
                selectStmt = firstParserRuleContext.GetDescendent<VBAParser.SelectCaseStmtContext>();
                var visitor = UnreachableCaseInspection.CreateParseTreeValueVisitor
                    (   ValueFactory, 
                        (ParserRuleContext context) =>
                            { return UnreachableCaseInspection.GetIdentifierReferenceForContext(context, state); }
                    );
                valueResults = selectStmt.Accept(visitor);
            }
            return valueResults;
        }
    }
}
