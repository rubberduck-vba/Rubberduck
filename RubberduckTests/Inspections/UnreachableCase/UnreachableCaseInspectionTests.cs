using System;
using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Moq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Refactoring.ParseTreeValue;

namespace RubberduckTests.Inspections.UnreachableCase
{
    [TestFixture]
    public class UnreachableCaseInspectionTests : InspectionTestsBase
    {
        [TestCase("Dim Hint$\r\nSelect Case Hint$", @"""Here"" To ""Eternity"",""Forever""", "String")] //String
        [TestCase("Dim Hint#\r\nHint#= 1.0\r\nSelect Case Hint#", "10.00 To 30.00, 20.00", "Double")] //Double
        [TestCase("Dim Hint!\r\nHint! = 1.0\r\nSelect Case Hint!", "10.00 To 30.00,20.00", "Single")] //Single
        [TestCase("Dim Hint%\r\nHint% = 1\r\nSelect Case Hint%", "10 To 30,20", "Integer")] //Integer
        [TestCase("Dim Hint&\r\nHint& = 1\r\nSelect Case Hint&", "1000 To 3000,2000", "Long")] //Long
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectExprTypeHint(string typeHintExpr, string firstCase, string expected)
        {
            string inputCode =
$@"
        Sub Foo()

        {typeHintExpr}
            Case {firstCase}
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
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
$@"
        Sub Foo(x As Variant)

        {typeHintExpr}

        Select Case x
            Case {firstCase}
            'OK
            Case Else
            'Unreachable
        End Select

        End Sub";
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
$@"
        Private Function ToLong(val As Variant) As Long
            ToLong = 5
        End Function

        Private Function ToString(val As Variant) As String
            ToString = ""Foo""
        End Function

        Sub Foo({argList})

            Select Case {selectExpr}
                Case 45
                'OK
                Case Else
                'OK
            End Select

        End Sub";
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
$@"
        Private Function ToLong(val As Variant) As Long
            ToLong = 5
        End Function

        Private Function ToString(val As Variant) As String
            ToString = ""Foo""
        End Function

        Sub Foo(x As Variant)

            Select Case x
                Case {rangeExpr1}, {rangeExpr2}
                'OK
                Case Else
                'OK
            End Select

        End Sub";
            var result = GetSelectExpressionType(inputCode);
            Assert.AreEqual(expected, result);
        }

        [TestCase("45", "55", "Integer")]
        [TestCase("45.6", "55", "Double")]
        [TestCase(@"""Test""", @"""55""", "String")]
        [TestCase("True", "y < 6", "Boolean")]
        [TestCase("#12/25/2018#", "#07/04/1776#", "Date")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeUnrecognizedSelectExpressionTypes(string rangeExpr1, string rangeExpr2, string expected)
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
$@"
        Sub Foo(x As Collection)
            Select Case x(3)
                Case {rangeExpr1}
                'OK
                Case {rangeExpr2}
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

        [TestCase("x.Item(2)", "True", false)]
        [TestCase("x.Item(2)", "True, False", true)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_CaseClauseTypeUnrecognizedCaseExpressionType2(string rangeExpr1, string rangeExpr2, bool triggersCaseElse)
        {
            string inputCode =
$@"
        Sub Foo(x As Collection)
            Select Case x(3)
                Case {rangeExpr1}
                'OK
                Case {rangeExpr2}
                'OK
                Case {rangeExpr2}
                'unreachable
                Case Else
                'Depends on flag
            End Select
        End Sub";
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
$@"
        Sub Foo(x As Variant)
            Select Case x
                Case {rangeExpr1}
                'OK
                Case {rangeExpr2}
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
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, overflow: 2);
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

        //For all values of 'y', one of prior case statements equals Case #3
        [TestCase("y < 55", "y > 30", "True")]
        [TestCase("y = 2", "y = 30", "False")]
        [TestCase("y <= 55", "y > 75", "y <= 55")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ComparablePredicates(string case1, string case2, string case3)
        {
            string inputCode =
$@"Sub Foo(x As Long, y As Double)

        Select Case x
           Case {case1}
            'OK
           Case {case2}
            'OK
           Case {case3}
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
$@"Sub Foo(x As Boolean)

        Private Const fromVal As Long = 500
        Private Const toVal As Long = 0

        Select Case x
           Case {caseClause}
            'OK
           Case True
            'Unreachable
        End Select

        End Sub";
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
$@"Sub Foo(x As Long)

                Select Case x
                    Case {complexClause1}
                    'OK
                    Case {complexClause2}
                    'Unreachable - detected by text compare of range clause(s)
                End Select

                End Sub";
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
        [TestCase(@"z To 1000, 95, ""TEST""", 1, 1)]
        [TestCase(@"1 To ""Forever"", 55000", 1, 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NumberRangeMixedTypes(string firstCase, int unreachableCount, int mismatchCount)
        {
            string inputCode =
$@"Sub Foo(x As Integer, z As String)

        Select Case x
            Case {firstCase}
            'Mismatch - unreachable
            Case 1 To 50
            'OK
            Case 45
            'Unreachable
        End Select

        End Sub";
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
            'Exceeds Integer values - overflow
            Case 10,11,12
            'OK
            Case 15, 40000
            'Overflow exception
            Case Is < 4
            'OK
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, overflow: 2);
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

        [TestCase("True", "Is <= False", 1)]
        [TestCase("Is >= True", "False", 2)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanRelationalOps(string firstCase, string secondCase, int expected)
        {
            string inputCode =
$@"Sub Foo( x As Boolean)

        Select Case x
            Case {firstCase}
            'OK
            Case {secondCase}
            'Unreachable
            Case 95
            'Unreachable
        End Select

        End Sub";
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
        'Unreachable
        Case Else
        'OK
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("False To True", 0, 1)] //firstCase is inherently unreachable - malformed
        [TestCase("Is < True", 0, 1)]   //firstCase is inherently unreachable
        [TestCase("Is > False", 0, 1)]   //firstCase is inherently unreachable
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionInherentlyUnreachable(string firstCase, int unreachableCount, int inherentlyUnreachable)
        {
            string inputCode =
$@"
Sub Foo(x As Boolean)
    Select Case x
        Case {firstCase} 
        'Reachable depends on firstCase
        Case True
        'OK
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachableCount, inherentlyUnreachable: inherentlyUnreachable);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("True To False", 1, 1)] //firstCase filters all possible values
        [TestCase("Is >= True", 1, 1)]   //firstCase filters all possible values
        [Category("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionAllValues(string firstCase, int unreachableCount, int caseElseCount)
        {
            string inputCode =
$@"
Private Function Random() As Double
    Random = VBA.Rnd()
End Function

Sub Foo(x As Boolean)
    Select Case Random() > 0.5
        Case {firstCase} 
        'Reachable depends on firstCase
        Case True
        'Reachable depends on firstCase
        Case Else
        'Reachable depends on firstCase
    End Select
End Sub";
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
            'Mismatch - inherently unreachable
            Case ""85""
            'OK
            Case 2
            'Unreachable
            Case 92
            'Unreachable
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2, mismatch: 1);
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
            Case 2 * 2 'Wednesday
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

        //https://github.com/rubberduck-vba/Rubberduck/issues/4119
        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_Enumeration()
        {
            const string inputCode =
@"
Option Explicit

Public Enum LogLevel
    TraceLevel = 0
    DebugLevel
    InfoLevel
    WarnLevel
    ErrorLevel
    FatalLevel
End Enum

Public Function LogLevelToString(ByVal level As LogLevel) As String
    Select Case level

        Case LogLevel.DebugLevel
            LogLevelToString = ""DEBUG""
        Case LogLevel.ErrorLevel
            LogLevelToString = ""ERROR""
        Case LogLevel.FatalLevel
            LogLevelToString = ""FATAL""
        Case LogLevel.InfoLevel
            LogLevelToString = ""INFO""
        Case LogLevel.TraceLevel
            LogLevelToString = ""TRACE""
        Case LogLevel.WarnLevel
            LogLevelToString = ""WARNING""
        Case 5 'Fatal Level
            'Unreachable - find something in order to test the test
    End Select
End Function";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumerationDefaultsNoAssignments()
        {
            const string inputCode =
@"
private Enum Fruit
    Apple
    Pear
    Orange
End Enum

Sub Foo(z As Fruit)
    Select Case z
        Case Apple
        'OK
        Case Pear 
        'OK     
        Case 1        
        'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumerationDefaultsNonIncremental()
        {
            const string inputCode =
@"
private Enum Fruit
    Apple = 0
    Pear = 10
    Orange
End Enum

Sub Foo(z As Fruit)
    Select Case z
        Case 11
        'OK
        Case Pear 
        'OK     
        Case Orange        
        'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumerationDefaultsRespectsPriorValues1()
        {
            const string inputCode =
@"
private Enum BitCountMaxValues
    max1Bit = 2 ^ 0
    max2Bits = 2 ^ 1 + max1Bit
    max3Bits
    max4Bits = 2 ^ 3 + max3Bits
End Enum

Sub Foo(z As BitCountMaxValues)
    Select Case z
        Case BitCountMaxValues.max3Bits
        'OK
        Case 4
        'Unreachable
    End Select
End Sub";
            var (expectedMsg, actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumerationDefaultsRespectsPriorValues2()
        {
            const string inputCode =
@"
private Enum BitCountMaxValues
    max1Bit
    max2Bits = 2 ^ 1 + max1Bit
    max3Bits
    max4Bits = 2 ^ 3 + max3Bits
End Enum

Sub Foo(z As BitCountMaxValues)
    Select Case z
        Case BitCountMaxValues.max3Bits
        'OK
        Case 3
        'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }


        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_EnumerationDefaultsRespectsPriorValues3()
        {
            const string inputCode =
@"
private Enum TestEnumVals
    Foo = -42
    PostFoo
    Bar = 9
    PostBar
    Baz = 0
    PostBaz
End Enum

Sub Foo(z As TestEnumVals)
    Select Case z
        Case TestEnumVals.PostFoo
        'OK
        Case -41
        'Unreachable
        Case TestEnumVals.PostBar
        'OK
        Case 10
        'Unreachable
        Case TestEnumVals.PostBaz
        'OK
        Case 1
        'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
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
$@"
        private Const maxValue As Long = 5000

        Sub Foo(z As Long)

        Select Case {selectExpr}
            Case {firstCase}
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
$@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case 95 
            'OK
            Case Is {opSymbol} x
            'OK
            Case -3 To 10 
            'OK - covers True and False
            Case Is {opSymbol} y 
            'OK
            Case Is {opSymbol} x 
            'Unreachable
            Case z < x
            'Unreachable
        End Select

        End Sub";
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
$@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case Is < x 
            'OK
            Case {secondCase}
            'OK
            Case {thirdCase}
            'Unreachable
        End Select

        End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("Eqv")]
        [TestCase("Imp")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ImpEqvOperatorsVariable(string op)
        {
            string inputCode =
$@"
        Sub Foo(x As Long, y As Long, z As Long)
        Select Case z
            Case Is < x 
            'OK
            Case x {op} y
            'OK
            Case -3 To 10 
            'OK
            Case x {op} y
            'Unreachable
            Case x {op} z
            'OK
        End Select

        End Sub";
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
End Sub";
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
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

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
$@"
Option Explicit

Private myVal As {propertyType}

Public Property Set AValue(val As {propertyType})
    myVal = val
End Property

Public Property Get AValue() As {propertyType}
    AValue = myVal
End Property
";
            var components = new List<(string moduleName, string inputCode, ComponentType componentType)>()
            {
                ("TestModule1", inputCode, ComponentType.StandardModule),
                ("Class1", inputClassCode, ComponentType.ClassModule)
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
$@"
Option Explicit

Private myVal As {propertyType}

Public Property Set AValue(val As {propertyType})
    myVal = val
End Property

Public Property Get AValue() As {propertyType}
    AValue = myVal
End Property
";
            var components = new List<(string moduleName, string inputCode, ComponentType componentType)>()
            {
                ("TestModule1",inputCode, ComponentType.StandardModule),
                ("Class1", inputClassCode, ComponentType.ClassModule)
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
$@"
Option Explicit

Public Const MY_CONSTANT As {propertyType} 
";
            var components = new List<(string moduleName, string inputCode, ComponentType componentType)>()
            {
                ("TestModule1",inputCode, ComponentType.StandardModule),
                ("TestModule2", inputModule2Code, ComponentType.StandardModule)
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
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase(@"x Like ", @"""*Bar""", 1)]
        [TestCase(@"y Like ", @"""*Bar""", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_Like(string caseClauseExpression, string likePattern, int expectedUnreachable)
        {
            string inputCode =
$@"
Private Const TEST_CONST As String = ""FooBar""

Sub FirstSub(x As String, y As String)
    Select Case True
        Case {caseClauseExpression}{likePattern}
            'OK
        Case TEST_CONST
            'Mismatch
        Case ""CandyBar"" Like {likePattern}
            'OK
        Case ""HandleBar"" Like {likePattern}
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: expectedUnreachable, mismatch: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LikeFilter()
        {
            string inputCode =
$@"
Private Const TEST_CONST As String = ""FooBar""

Sub FirstSub(x As String, y As String)
    Select Case True
        Case y Like ""*Bar""
            'OK
        Case y Like ""*""
            'OK
        Case y Like ""aaaa??""
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, mismatch: 0);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase(@"Option Compare Binary", @"""f"" Like ""[a-z]*""", 2)]
        [TestCase(@"", @"""f"" Like ""[a-z]*""", 2)]
        [TestCase(@"Option Compare Binary", @"""F"" Like ""[a-z]*""", 2)]
        [TestCase(@"", @"""F"" Like ""[a-z]*""", 2)]
        [TestCase(@"Option Compare Text", @"""f"" Like ""[a-z]*""", 2)]
        [TestCase(@"Option Compare Text", @"""F"" Like ""[a-z]*""", 2)]
        [TestCase(@"Option Compare Database", @"""f"" Like ""[a-z]*""", 2)]
        [TestCase(@"Option Compare Database", @"""F"" Like ""[a-z]*""", 2)]
        [TestCase(@"", @"x Like ""*""", 2)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LikeRespectsOptionCompareSetting( string optionCompare, string case1, int expectedUnreachable)
        {
            string inputCode =
$@"
Option Explicit
{optionCompare}

Sub FirstSub(x As String)
    Select Case True
        Case {case1}
            'Unreachable - if case1 is false
        Case 5 < 6
            'Unreachable - if case1 is True
        Case 9 > 8
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: expectedUnreachable);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_LikeSpecialCaseAsterisk()
        {
            string inputCode =
@"
Sub FirstSub(x As String, y As String, flag As Boolean)
    Select Case flag
        Case x Like ""*""
            'OK
        Case Not x Like ""*""
            'OK
        Case y Like ""???*""
            'Unreachable
        Case Else
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("Select Case 100", @"""100""", 2)]
        [TestCase("Select Case True", @"x Like ""*""", 2)]
        [TestCase("Select Case False", @"Not x Like ""*""", 2)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectExpressionIsAConstant( string selectCase, string case1, int unreachableCount)
        {
            string inputCode =
$@"
Sub FirstSub(x As String, y As String)
    {selectCase}
        Case {case1}
            'OK
        Case y Like ""*""
            'Unreachable
        Case y Like ""???*""
            'Unreachable
        Case Else
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachableCount, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase(@"Option Compare Binary", @"""A"" > ""a""", 2)]
        [TestCase(@"Option Compare Text", @"""A"" > ""a""", 1)]
        [TestCase(@"Option Compare Binary", @"""A"" = ""a""", 1)]
        [TestCase(@"Option Compare Text", @"""A"" = ""a""", 2)]
        [TestCase(@"", @"""A"" = ""a""", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_StringCompareRespectsOptionCompareSetting(string optionCompare, string case1, int unreachable)
        {
            string inputCode =
$@"
{optionCompare}

Sub FirstSub(y As Boolean)

Select Case y
        Case {case1}
            'OK
        Case 10 > 2
            'Unreachable
        Case 7 < 10
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachable);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase(@"Option Compare Binary", 1)]
        [TestCase(@"Option Compare Text", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_StringCompares(string optionCompare, int unreachable)
        {
            string valueWithIgnorableChars = "\"Ani\u00ADmal\"";
            string inputCode =
$@"
{optionCompare}

Sub FirstSub(y As Boolean)

Const s1 As String = ""animal""
Const s2 As String = {valueWithIgnorableChars}

Select Case y
        Case s1 = s2    'Always false since we do not ignore the hyphen (""\u00AD"")
            'OK
        Case 10 > 2
            'OK because first case is always false
        Case 7 < 10
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachable);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_Ampersand()
        {
            string inputCode =
@"
Private Const TEST_CONST As String = ""Bar""

Sub FirstSub(x As String)
    Select Case x
        Case ""Foo"" & TEST_CONST
            'OK
        Case ""FooBar""
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_AmpersandWithVariable()
        {
            string inputCode =
@"
Sub FirstSub(x As String, y As Long)
Private Const TEST_CONST As String = ""o""

Select Case x
        Case ""Foo"" & y & ""Bar""
            'OK
        'Case ""Foo1""
            'OK
        Case ""Fo"" & TEST_CONST & y & ""Bar""
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_AmpersandMismatch()
        {
            string inputCode =
@"
Sub FirstSub(x As String, y As Long)

Select Case y
        Case 45 & ""B""
            'Mismatch
        Case 45 + 2
            'OK
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 0, mismatch: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/3962
        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_AdditionString()
        {
            //Resolving ""2"" + ""1"" to 21 (correct) yields 3 unreachable cases
            //Resolving ""2"" + ""1"" to 3 (incorrect) would yield 1 unreachable case
            string inputCode =
@"
Sub FirstSub(bar As Double)

    Select Case bar
        Case ""2"" + ""1""
        Case 21
            'Unreachable
        Case 3 * 7
            'Unreachable
        Case (""1"" + ""0"") * 3 - 9 
            'Unreachable
        Case 3
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("As Double")]
        [TestCase("As Long")]
        [TestCase("As Byte")]
        [TestCase("As Currency")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SelectExpressionConstant(string typeName)
        {
            string inputCode =
$@"
private Const AVALUE {typeName} = 15

Sub FirstSub()

    Dim bar {typeName}
    bar = 15

    Select Case AVALUE
        Case bar
            'Unreachable
        Case 22
            'Unreachable
        Case 89
            'Unreachable
        Case 0 To 10
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("1# * .00125#", 3, 1)]
        [TestCase("1@ * .0012@", 3, 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_Currency(string thirdCase, int unreachable, int caseElse)
        {
            string inputCode =
$@"
Sub FirstSub()

        Const currencyVal As Currency = 1 * .0012@
        Const doubleVal As Double = 1 * .00125#

        Select Case currencyVal
            Case doubleVal  
                'OK
            Case doubleVal * 1
                'Unreachable
            Case {thirdCase}
                'Unreachable
            Case 0.25
                'Unreachable - Select Case value is a constant
            Case Else
                'Unreachable
        End Select
End Sub";

            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: unreachable, caseElse: caseElse);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("#12/1/2020#,#12/2/2020#", "#12/1/2020#")]
        [TestCase("Is < #12/1/2020#", "Is < #12/1/2010#")]
        [TestCase("#1/1/2020# To #8/1/2020#", "#7/1/2020#")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_DateType(string case1, string case2)
        {
            string inputCode =
$@"
Sub FirstSub(bar As Date)

    Select Case bar
        Case {case1}
            'OK
        Case {case2}
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("Is > #1/1/2020#", "Is < #7/1/2020#")]
        [TestCase("Is < #1/1/2020#, Is > #1/2/2020#", "#1/1/2020#,#1/2/2020#")]
        [TestCase("Is < #1/1/2020#, Is > #10/2/2020#", "#1/1/2019# To #9/2/2035#")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_DateTypeCoversAll(string case1, string case2)
        {
            string inputCode =
$@"
Sub FirstSub(bar As Date)

    Select Case bar
        Case {case1}
            'OK
        Case {case2}
            'OK
        Case Else
            'Unreachable
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [TestCase("AVALUE + ANOTHERVALUE")]
        [TestCase("AVALUE * ANOTHERVALUE")]
        [TestCase("AVALUE + ANOTHERVALUE To 700")]
        [TestCase("Is < AVALUE + ANOTHERVALUE")]
        [TestCase("x < AVALUE + ANOTHERVALUE")]
        [TestCase("-(AVALUE + ANOTHERVALUE)")]
        [TestCase("2147483678")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_Overflow(string firstCase)
        {
            string inputCode =
$@"
private Const AVALUE As Byte = 250
private Const ANOTHERVALUE As Byte = 250

Sub FirstSub(x As Integer)
    Select Case x
        Case {firstCase}
            'Overflow
        Case 20 + 2
            'OK
        Case ""2"" + ""2""
            'Unreachable
        Case AVALUE + ""7""
            'OK
    End Select
End Sub";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 1, overflow: 1);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_TypeMismatch()
        {
            string inputCode =
@"
Private Sub Foo(x As Date)
    Select Case x
        Case ""Hoosier Daddy""
            MsgBox ""Hoosier Daddy""    'mismatch - found during inspection
        Case ""Test"" And ""Check""
            MsgBox ""'Test' And 'Check'""   'mismatch - found while parsing
        Case ""1/1/2020""
            MsgBox ""1/1/2020 triggered""
    End Select
End Sub
";
            (string expectedMsg, string actualMsg) = CheckActualResultsEqualsExpected(inputCode, unreachable: 0, mismatch: 2);
            Assert.AreEqual(expectedMsg, actualMsg);
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_VbObjectErrorConstant()
        {
            var expectedUnreachableCount = 2;
            string inputCode =
@"
Enum Fubar
    Foo = vbObjectError + 1
    Bar = vbObjectError + 2
End Enum

Sub Example(value As Long)
    Select Case value
        Case Fubar.Foo
            Debug.Print ""Foo""
        Case Fubar.Bar
            Debug.Print ""Bar""
        Case vbObjectError + 1 'unreachable
            Debug.Print ""Unreachable""
        Case -2147221502 'unreachable
            Debug.Print ""Unreachable""
    End Select
End Sub
";

            var vbe = CreateStandardModuleProject(inputCode);

            IEnumerable<IInspectionResult> actualResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = InspectionUnderTest(state, TestGetValuedDeclaration);

                WalkTrees(inspection, state);
                actualResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Unreachable));

            Assert.AreEqual(expectedUnreachableCount, actualUnreachable.Count());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/4680
        [TestCase("vbNewLine", "vbCr + vbLf")]
        [TestCase("vbNewLine", "Chr(13) + Chr(10)")]
        [TestCase("vbNewLine", "Chr$(13) + Chr$(10)")]
        [TestCase("Chr(13) + Chr(10)", "Chr$(13) + Chr$(10)")]
        [TestCase("vbCr + vbLf", "vbNewLine")]
        [TestCase("vbCr + Chr(10)", "vbNewLine")]
        [TestCase("Chr(13) + vbLf", "vbNewLine")]
        [TestCase("Chr(0)", "vbNullChar")]
        [TestCase("Chr$(0)", "vbNullChar")]
        [TestCase("Chr(8)", "vbBack")]
        [TestCase("Chr$(8)", "vbBack")]
        [TestCase("Chr(12)", "vbFormFeed")]
        [TestCase("Chr$(12)", "vbFormFeed")]
        [TestCase("Chr(9)", "vbTab")]
        [TestCase("Chr$(9)", "vbTab")]
        [TestCase("Chr(11)", "vbVerticalTab")]
        [TestCase("Chr$(11)", "vbVerticalTab")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NonPrintingControlConstants(string testCase, string equivalent)
        {
            var expectedUnreachableCount = 1;
            string inputCode =
$@"
Sub Foo(value As String)
    Select Case value
        Case ""Hello"" + {testCase} + ""World""
            MsgBox ""testCase version""
        Case ""Hello"" + {equivalent} + ""World"" 'unreachable
            MsgBox ""equivalent version""
        Case ""Reachable""
            MsgBox ""Reachable""
    End Select
End Sub
";
            var vbe = CreateStandardModuleProject(inputCode);

            IEnumerable<IInspectionResult> actualResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = InspectionUnderTest(state, TestGetValuedDeclaration);

                WalkTrees(inspection, state);
                actualResults = inspection.GetInspectionResults(CancellationToken.None);
            }

            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Unreachable));

            Assert.AreEqual(expectedUnreachableCount, actualUnreachable.Count());
        }

        private static Dictionary<string, (string, string)> _vbConstConversions = new Dictionary<string, (string, string)>()
        {
            ["vbNewLine"] = ("Chr$(13) & Chr$(10)", Tokens.String),
            ["vbCr"] = ("Chr$(13)", Tokens.String),
            ["vbLf"] = ("Chr$(10)", Tokens.String),
            ["vbNullChar"] = ("Chr$(0)", Tokens.String),
            ["vbBack"] = ("Chr$(8)", Tokens.String),
            ["vbTab"] = ("Chr$(9)", Tokens.String),
            ["vbVerticalTab"] = ("Chr$(11)", Tokens.String),
            ["vbFormFeed"] = ("Chr$(12)", Tokens.String),
            ["vbObjectError"] = ("-2147221504", Tokens.Long),
        };

        private static (bool IsType, string ExpressionValue, string TypeName) TestGetValuedDeclaration(Declaration declaration)
        {
            if (!_vbConstConversions.ContainsKey(declaration.IdentifierName))
            {
                return (false, null, null);
            }

            (string expressionValue, string typename) = _vbConstConversions[declaration.IdentifierName];
            return (true, expressionValue , typename);
        }

        private (string expectedMsg, string actualMsg) CheckActualResultsEqualsExpected(string inputCode, int unreachable = 0, int mismatch = 0, int caseElse = 0, int inherentlyUnreachable = 0, int overflow = 0)
        {
            var components = new List<(string moduleName, string inputCode, ComponentType componentType)>() { ("TestModule1", inputCode, ComponentType.StandardModule) };
            return CheckActualResultsEqualsExpected(components, unreachable, mismatch, caseElse, inherentlyUnreachable, overflow);
        }

        private (string expectedMsg, string actualMsg) CheckActualResultsEqualsExpected(List<(string moduleName, string inputCode, ComponentType componentType)> components, int unreachable = 0, int mismatch = 0, int caseElse = 0, int inherentlyUnreachable = 0, int overflow = 0)
        {
            var expected = new Dictionary<string, int>
            {
                { Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Unreachable, unreachable },
                { Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_InherentlyUnreachable, inherentlyUnreachable },
                { Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_TypeMismatch, mismatch },
                { Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Overflow, overflow },
                { Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_CaseElse, caseElse },
            };

            var actualResults = InspectionResultsForModules(components).ToList();

            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Unreachable));
            var actualMismatches = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_TypeMismatch));
            var actualUnreachableCaseElses = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_CaseElse));
            var actualInherentUnreachable = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_InherentlyUnreachable));
            var actualOverflow = actualResults.Where(ar => ar.Description.Equals(Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Overflow));

            var actualMsg = BuildResultString(actualUnreachable.Count(), actualMismatches.Count(), actualUnreachableCaseElses.Count(), actualInherentUnreachable.Count(), actualOverflow.Count());
            var expectedMsg = BuildResultString(expected[Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Unreachable], 
                expected[Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_TypeMismatch], 
                expected[Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_CaseElse],
                expected[Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_InherentlyUnreachable],
                expected[Rubberduck.Resources.Inspections.InspectionResults.UnreachableCaseInspection_Overflow]
                );

            return (expectedMsg, actualMsg);
        }

        private Mock<IVBE> CreateStandardModuleProject(string inputCode)
            => MockVbeBuilder.BuildFromModules(new List<(string moduleName, string inputCode, ComponentType componentType)>() { ("TestModule1", inputCode, ComponentType.StandardModule) });

        private static string BuildResultString(int unreachableCount, int mismatchCount, int caseElseCount, int inherentCount, int overflowCount)
            => $"Unreachable={unreachableCount}, Mismatch={mismatchCount}, CaseElse={caseElseCount}, Inherent={inherentCount}, Overflow={overflowCount}";

        private string GetSelectExpressionType(string inputCode)
        {
            var selectStmtValueResults = GetParseTreeValueResults(inputCode, out VBAParser.SelectCaseStmtContext selectStmtContext, out var module);

            var inspector = TestUnreachableCaseInspector();
            return inspector.SelectExpressionTypeName(selectStmtContext, selectStmtValueResults);
        }

        private IParseTreeVisitorResults GetParseTreeValueResults(string inputCode, out VBAParser.SelectCaseStmtContext selectStmt, out QualifiedModuleName contextModule)
        {
            selectStmt = null;
            IParseTreeVisitorResults valueResults;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var finder = state.DeclarationFinder;
                var (parseTreeModule, moduleParseTree) = state.ParseTrees
                    .First(pt => pt.Value is ParserRuleContext);
                selectStmt = ((ParserRuleContext)moduleParseTree).GetDescendent<VBAParser.SelectCaseStmtContext>();
                var visitor = TestParseTreeValueVisitor();
                valueResults = visitor.VisitChildren(parseTreeModule, selectStmt, finder);
                contextModule = parseTreeModule;
            }
            return valueResults;
        }

        private IParseTreeValueVisitor TestParseTreeValueVisitor(Func<Declaration, (bool, string, string)> valueDeclarationEvaluator = null)
        {
            var valueFactory = new ParseTreeValueFactory();
            return new ParseTreeValueVisitor(valueFactory, valueDeclarationEvaluator);
        }

        private IUnreachableCaseInspector TestUnreachableCaseInspector()
        {
            var valueFactory = new ParseTreeValueFactory();
            return new UnreachableCaseInspector(valueFactory);
        }

        private IParseTreeInspection InspectionUnderTest(RubberduckParserState state, Func<Declaration, (bool, string, string)> valueDeclarationEvaluator)
        {
            var inspector = TestUnreachableCaseInspector();
            var parseTeeValueVisitor = TestParseTreeValueVisitor(valueDeclarationEvaluator);
            return new UnreachableCaseInspection(state, inspector, parseTeeValueVisitor);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            var inspector = TestUnreachableCaseInspector(); 
            var parseTeeValueVisitor = TestParseTreeValueVisitor();
            return new UnreachableCaseInspection(state, inspector, parseTeeValueVisitor);
        }
    }
}
