using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class UnreachableCaseInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SingleUnreachableCase()
        {
            const string inputCode =
@"Sub Foo()

Const x As String =""Bar""
Select Case x
  Case ""Foo"", ""Bar""
    'OK
  Case ""Food""
    'OK
  Case ""Bar""
    'Unreachable
  Case ""Foodie""
    'OK
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_RangeConflictsWithPriorIsStmt()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case Is > 5
    'OK
  Case 4 To 10
    'OK - overlap
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_DuplicateComplexCaseClause()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case Is > 500
    'OK
  Case x^2 = 49, CLng(VBA.Rnd() * 100) * x < 30
    'OK
  Case 45 to 100
    'OK
  Case x^2 = 49, CLng(VBA.Rnd() * 100) * x < 30
    'Unreachable - Copy/Paste
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_DuplicateComplexCaseClauseOutofOrder()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case x^2 = 49, CLng(VBA.Rnd() * 100) * x < 30
    'OK
  Case 45 to 100
    'OK
  Case CLng(VBA.Rnd() * 100) * x < 30, x^2 = 49
    'Unreachable - Copy/Paste
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_InternalOverlap()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case Is > 5, 15, 20, Is < 55
    'Conflict
  Case 4 To 10
    'Unreachable
  Case True
    'Unreachable
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_MultiplyUsedSelectVariable()
        {
            const string inputCode =
@"Sub Foo(caseNum As Long)

Const x As String =""Bar""
Select Case x
  Case ""Foo"", ""Bar""
    'Do FooBar
  Case ""Foo""
    'Unreachable
  Case ""Bar""
    'Unreachable
  Case Else
    MsgBox ""Unable To handle "" & x 
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_MultipleUnreachableCases()
        {
            const string inputCode =
@"Sub Foo()

Const x As String =""Bar""
Select Case x
  Case ""Foo"", ""Bar""
    'Do FooBar
  Case ""Foo""
    'Unreachable
  Case ""Bar""
    'Unreachable
  Case ""Foo""
    'Unreachable
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ConflictingCaseStmt()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 1 To 45, 35, 85
    'Internal Conflict but reachable
  Case Else
    'OK
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_OverlapOnlyCaseStmt()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 40
    'OK
  Case 1 To 45, 35, 85
    'Internal Conflict but reachable
  Case Else
    'OK
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CoverAllCases()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case x > -5000
    'OK
  Case Is < 5
    'Conflict
  Case 500 To 700
    'Unreachable
  Case Else
    'Unreachable
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CoverAllCasesSingleClause()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case  -5000, Is <> -5000
    'OK
  Case Is > 5
    'Unreachable
  Case 500 To 700
    'Unreachable
  Case Else
    'Unreachable
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CaseStmtCoversAll()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 1 To 45, Is < 5, x > -5000
    'Internal Conflict but reachable
  Case 5500
    'Unreachable
  Case 500 To 700
    'Unreachable
  Case Else
    'Unreachable
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CaseStmtCoversAllNEQ()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case x = -4, x <> -4
    'Only reachable case
  Case 5500
    'Unreachable
  Case 500 To 700
    'Unreachable
  Case Else
    'Unreachable
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_StringRange()
        {
            const string inputCode =
@"Sub Foo()

Const x As String =""Bar""
Select Case x
  Case ""Alpha"" To ""Omega""
    'Do FooBar
  Case ""Alphabet""
    'Unreachable
  Case ""Ohm""
    'Unreachable
  Case ""Omegaad""
    'OK
End Select
                
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ReusedSelectExpressionVariable()
        {
            const string inputCode =
@"Sub Foo()

Const x As String =""Bar""
  
Select Case x
  Case ""Foo"", ""Bar""
    'Do FooBar
  Case ""Foo""
    'Unreachable
  Case ""Food""
        'OK
End Select
                
End Sub";

            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRange()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 1 To 100
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRange2()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 150 To 250
    'Do FooBar1
  Case 1 To 100
    'Do FooBar2
  Case 101 To 149
    'Do FooBar2
  Case 25 To 249 
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeHighToLow()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 100 To 1
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeAndIsStmt()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 1 To 100
    'Do FooBar
  Case Is <= 150
    'Conflict
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeInverted()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 100 To 1
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeConflict()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 1 To 100
    'Do FooBar
  Case 101 To 125
    'reachable
  Case 50 To 200
    'Conflict
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeFollowsScalar()
        {
            const string inputCode =
@"Sub Foo(x as Long)

Select Case x
  Case 55
    'OK
  Case 50 To 200
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_EmbeddedSelectCase()
        {
            const string inputCode =
@"Sub Foo(x As Long, z As Long) 

Select Case x
  Case 1 To 10
    'Do FooBar
  Case 9
    'Unreachable
  Case 11
    Select Case  z
      Case 5 To 25
        'Do FooBar
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_EmbeddedSelectCaseStringType()
        {
            const string inputCode =
@"Sub Foo()

Const x As String = ""Foo""
Const z As String = ""Bar""

Select Case x
  Case ""Foo"", ""Bar""
    'Do FooBar
  Case ""Foo""
    'Unreachable
  Case ""Food""
    Select Case  z
      Case ""Foo"", ""Bar"",""Goo""
        'Do FooBar
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SimpleLongCollision()
        {
            const string inputCode =
@"Sub Foo(x As Long)
Select Case x
  Case 1,2,-5
    'Do FooBar
  Case 2
    'Unreachable
  Case -5
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ExceedsIntegerValue()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

Select Case x
  Case 10,11,12
    'Do FooBar
  Case 40000
    'Exceeds Integer value
  Case x < 4
    'OK
  Case -50000
    'Exceeds Integer values
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, outOfRange: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ExceedsLongValue()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
 ' Case 98
    'OK
  Case 5 To 25, 50, 80
    'OK
  Case 214#
    'OK
  Case 2147486648#
    'Unreachable
  Case -2147486649#
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, outOfRange: 2);
        }


        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IntegerWithDoubleValue()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

Select Case x
  Case 214.0
    'OK - ish
  Case -214#
    'OK - ish
  Case Is < -5000
    'OK
  Case 98
    'OK
  Case 5 To 25, 50, 80
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ExceedsCurrencyValue()
        {
            const string inputCode =
@"Sub Foo(x As Currency)

Select Case x
  Case 85.5
    'Do FooBar
  Case -922337203685477.5809
    'Unreachable
  Case 922337203685490.5808
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, outOfRange: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ExceedsSingleValue()
        {
            const string inputCode =
@"Sub Foo(x As Single)

Select Case x
  Case 85.5
    'Do FooBar
  Case -3402824E38
    'Unreachable
  Case 3402824E38
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, outOfRange: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumbersAsBooleanCases()
        {
            const string inputCode =
@"Sub Foo(x As Boolean)

Select Case x
  Case -5
    'Evaluates as 'True'
  Case 4
    'Unreachable
  Case 1
    'Unreachable
  Case Else
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_HandlesBooleanConstants()
        {
            const string inputCode =
@"Sub Foo(x As Boolean)

Select Case x
  Case True
    'OK
  Case False
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanUnreachableCaseElse()
        {
            const string inputCode =
@"Sub Foo(x As Boolean)

Select Case x
  Case -7
    'OK
  Case False
    'OK
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElseUsesLong()
        {
            const string inputCode =
@"Sub Foo(x As Boolean)

Select Case VBA.Rnd() > 0.5
  Case -7
    'OK
  Case False
    'OK
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElseUsingBoolean()
        {
            const string inputCode =
@"Sub Foo(x As Boolean)

Select Case VBA.Rnd() > 0.5
  Case True To False
    'OK
  Case True
    'Unreachable
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElseInvertBooleanRange()
        {
            const string inputCode =
@"Sub Foo(x As Boolean)

Select Case VBA.Rnd() > 0.5
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ExceedsByteValue()
        {
            const string inputCode =
@"Sub Foo(x As Byte)

Select Case x
  Case 2,5,7
    'Do FooBar
  Case 254
    'OK
  Case 255
    'OK
  Case 256
    'Out of Range
  Case -1
    'Out of Range
  Case 0
    'OK
  Case 1
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, outOfRange: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SimpleDoubleCollision()
        {
            const string inputCode =
@"Sub Foo(x As Double)

Select Case x
  Case 4.4, 4.5, 4.6
    'Do FooBar
  Case 4.5
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SimpleIntegerCollision()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

Select Case x
  Case 1,2,5
    'Do FooBar
  Case 2
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_StringWhereLongShouldBeIncludeLongAsString()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 1 To 49
    'OK
  Case ""51""
    'Reachable - VBA is happy to change this to 51
  Case ""Hello World""
    'Unreachable
  Case 50
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_DoubleWhereLongShouldBe()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 88.55
    'OK - but may not get what you expect
  Case Else
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_MultipleRanges()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 1 To 4, 7 To 9, 11, 13, 15 To 20
    'Do FooBar
  Case 8
    'Unreachable
  Case 11
    'Unreachable
  Case 17
    'Unreachable
  Case 21
    'Reachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CascadingIsStatements()
        {
            const string inputCode =
@"Sub Foo(LNumber As Long)

Select Case LNumber
   Case Is < 100
      LRegionName = ""North""
   Case Is < 200
      'Conflict  
      LRegionName = ""South""
   Case Is < 300
      LRegionName = ""East""
      'Conflict  
   Case Else
      LRegionName = ""West""
   End Select
End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CascadingIsStatementsGT()
        {
            const string inputCode =
@"Sub Foo()

Const LNumber As Long = 340

Select Case LNumber
   Case Is > 300
      LRegionName = ""North""
   Case Is > 200
      'Conflict  
      LRegionName = ""South""
   Case Is > 100
      LRegionName = ""East""
       'Conflict  
  Case Else
      LRegionName = ""West""
   End Select
End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGT()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is > 5000
    'Do FooBar
  Case 5000
    'Do Foobar again
  Case 5001
    'Unreachable
  Case 10000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmt()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is > 5000 
    'Do FooBar
  Case Is < 10000
    'Conflict
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachable()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is > 5000 
    'Do FooBar
  Case Is < 10000
    'Conflict
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachableLTE()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is > 5000 
    'Do FooBar
  Case Is <= 10000
    'Conflict
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachableUsingIs()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case z <> 5 
    'Do FooBar
  Case Is = 5
    'OK
  Case 400
    'Unreachable
  Case Else
    'Unreachable
End Select
End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1,  caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_Evil()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 1 To 4, 7 To 9, 11, 13, Is > 400 
    'Do FooBar
  Case 401
    'Unreachable
  Case 15 To 25, 300 To 401
    'Conflict
  Case z < 2
    'Conflict
  Case 88
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_RelationalOpSimple()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case z > 5000
    'Do FooBar
  Case 5000
    'Do Foobar again
  Case 5001
    'Unreachable
  Case 10000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseHasMultOp()
        {
            const string inputCode =
@"
Function Bar() As Long
    Bar = 5
End Function

Sub Foo(z As Long)

Select Case Bar()  * z
  Case Is > 5000
    'Do FooBar
  Case 5000
    'Do Foobar again
  Case 5001
    'Unreachable
  Case 10000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstant()
        {
            const string inputCode =
@"
private Const maxValue As Long = 5000

Sub Foo(z As Long)

Select Case z
  Case Is > maxValue
    'Do Foobar again
  Case 6000
    'Unreachable
  Case 8500
    'Unreachable
  Case 15
    'OK
  Case Else
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstantReferenceExpr()
        {
            const string inputCode =
@"
private Const maxValue As Long = 5000

Sub Foo(z As Long)

Select Case z
  Case z > maxValue
    'Do Foobar again
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstantReferenceOnRHS()
        {
            const string inputCode =
@"
private Const maxValue As Long = 5000

Sub Foo(z As Long)

Select Case z
  Case maxValue < z
    'Do Foobar again
  Case 6000
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstantIndeterminantExpression()
        {
            const string inputCode =
@"
private Const maxValue As Long = 5000

Sub Foo(z As Long)

Select Case z
  Case z > maxValue/2
    'OK
  Case z > maxValue
    'OK
  Case 15
    'OK
  Case 8500
    'Unreachable
  Case Else
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
    'Do FooBar
  Case 5000
    'Do Foobar again
  Case 5001
    'Unreachable
  Case 10000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_RelationalOpExpression()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 500 < z
    'Do FooBar
  Case 500
    'Do Foobar again
  Case 501
    'Unreachable
  Case 1000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ConstantSelectStmt()
        {
            const string inputCode =
@"Sub Foo()

Dim x As Long
x = 7

Dim y As Long
y = 9

Select Case True
  Case x >= 7
    'Do FooBar
  Case y >= 4
    'Do Foobar again
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGTFollowsRange()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 3 To 10
    'Do FooBar
  Case Is > 8
    'Conflict
  Case 4
    'Unreachable
  Case 2
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGTConflicts()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 5000
    'Do Foobar again
  Case 10000
    'OK
  Case Is > 5000
    'Conflict
  Case 5001
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGTE()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is >= 5000
    'Do FooBar
  Case 4999
    'Do Foobar again
  Case 5000
    'Unreachable
  Case 10000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtLT()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is < 5000
    'Do FooBar
  Case 7
    'Unreachable
  Case 4999
    'Unreachable
  Case 5000
    'Do Foobar again
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtLTE()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is <= 5000
    'Do FooBar
  Case 7
    'Unreachable
  Case 5001
    'Do Foobar again
  Case 5000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtLTEReverse()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 7
    'OK
  Case Is <= 5000
    'Conflict
  Case 5001
    'Do Foobar again
  Case 5000
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtNEQ()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is <> 8
    'Do FooBar
  Case 7
    'Unreachable
  Case 5001
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtNEQReverseOrder()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 7
    'OK
  Case Is <> 8
    'OK
  Case 5001
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CaseElseIsStmtNEQAndSingleValue()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 8
    'OK
  Case Is <> 8
    'OK
  Case -4000
    'Unreachable
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtNEQAllValues()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is <> 8
    'OK
  Case 8
    'OK
  Case 5001
    'Unreachable
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtEQ()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is = 8
    'OK
  Case 8
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtEQReveserOrder()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 8
    'OK
  Case Is = 8
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtMultiple()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is > 8
    'Do FooBar
  Case 8
    'OK
  Case  Is = 9
    'Unreachable
  Case Is < 100
    'OK
  Case Is < 5
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable:2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtAndRange()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is > 8
    'OK
  Case 3 To 10
    'Conflict - OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtAndNegativeRange()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is < 8
    'Do FooBar
  Case -10 To -3
    'Unreachable
  Case 0
    'Unreachable
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtAndRangeAreNegative()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case Is < -8
    'OK
  Case -10 To -3
    'Overlap - but reachable
  Case 0
    'OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SingleValueFollowedByIsStmt()
        {
            const string inputCode =
@"Sub Foo(z As Long)

Select Case z
  Case 200
    'OK
  Case Is > 8
    'Conflict but OK
End Select

End Sub";
            CheckActualResultsEqualsExpected(inputCode);
        }

        private void CheckActualResultsEqualsExpected(string inputCode, int unreachable = 0, int mismatch = 0, int outOfRange = 0, int caseElse = 0)
        {
            var expected = new Dictionary<string, int>
            {
                { CaseInspectionMessages.Unreachable, unreachable },
                { CaseInspectionMessages.MismatchType, mismatch },
                { CaseInspectionMessages.ExceedsBoundary, outOfRange },
                { CaseInspectionMessages.CaseElse, caseElse },
            };

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new SelectCaseInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(CaseInspectionMessages.Unreachable));
            var actualMismatches = actualResults.Where(ar => ar.Description.Equals(CaseInspectionMessages.MismatchType));
            var actualOutOfRange = actualResults.Where(ar => ar.Description.Equals(CaseInspectionMessages.ExceedsBoundary));
            var actualUnreachableCaseElses = actualResults.Where(ar => ar.Description.Equals(CaseInspectionMessages.CaseElse));

            Assert.AreEqual(expected[CaseInspectionMessages.Unreachable], actualUnreachable.Count(), "Unreachable result");
            Assert.AreEqual(expected[CaseInspectionMessages.MismatchType], actualMismatches.Count(), "Mismatch result");
            Assert.AreEqual(expected[CaseInspectionMessages.ExceedsBoundary], actualOutOfRange.Count(), "Boundary Check result");
            Assert.AreEqual(expected[CaseInspectionMessages.CaseElse], actualUnreachableCaseElses.Count(), "CaseElse result");
        }
    }
}
