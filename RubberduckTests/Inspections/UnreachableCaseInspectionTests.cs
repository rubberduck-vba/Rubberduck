using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

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
@"Sub Foo(caseNum As Long)

Const x As String =""Bar""
Select Case x
  Case ""Foo"", ""Bar""
    'Do FooBar
  Case ""Food""
    'Unreachable
  Case ""Bar""
    'Unreachable
  Case ""Foodie""
    'Unreachable
End Select
                
End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 3);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
  Case ""Bar""
    'Unreachable
  Case ""Foo""
    'Unreachable
End Select
                
Select Case x
  Case ""Foo"", ""Bar""
    'Do FooBar
  Case ""Foo""
    'Unreachable
  Case ""Bar""
    'Unreachable
  Case ""Food""
End Select
                
End Sub

Private Function AppendBar(x As String) As String
    AppendBar = x + ""Bar""
End Function";

            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 5);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRange()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 1 To 100
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRange2()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 100 To 1
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRange3()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 1 To 100
    'Do FooBar
  Case Is <= 150
    'Conflict
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeInverted()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 100 To 1
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeConflict()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 1 To 100
    'Do FooBar
  Case 101 To 125
    'reachable
  Case 50 To 200
    'Unreachable - conflict
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_NumberRangeFollowsScalar()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 55
    'Do FooBar
  Case 50 To 200
    'Unreachable - conflict
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_EmbeddedSelectCase()
        {
            const string inputCode =
@"Sub Foo()

Const x As Long = 7
Const z As Long = 5

Select Case x
  Case 1 To 10
    'Do FooBar
  Case 9
    'Unreachable - conflict
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 4);
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
    'Unreachable - conflict
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 4);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
  'Case 40000
  '  'Unreachable
  Case x < 4
    'OK
  'Case -50000
  '  'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }


        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IntegerWithDoubleValue()
        {
            const string inputCode =
@"Sub Foo(x As Integer)

Select Case x
  Case 214.0
    'Unreachable
  Case -214#
    'Unreachable
  Case Is < -5000
    'OK
  Case 98
    'OK
  Case 5 To 25, 50, 80
    'OK
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElse()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElse2()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElse3()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_ExceedsByteValue()
        {
            const string inputCode =
@"Sub Foo(x As Byte)

Select Case x
  Case 1,2,5
    'Do FooBar
  Case 256
    'Unreachable
  Case -1
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_StringWhereLongShouldBe()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
'  Case 1 To 49
    'Do FooBar
'  Case 50
    'Reachable
  Case ""Test""
    'Unreachable
  Case ""85""
    'OK
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_StringWhereLongShouldBe2()
        {
            const string inputCode =
@"Sub Foo(x As Long)

Select Case x
  Case 1 To 49
    'Do FooBar
  Case ""51""
    'Reachable - VBA is happy to change this to 51
  Case ""Hello World""
  Case 50
    'Reachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 3);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_CascadingIsStatements2()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGT()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmt()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is > 5000 
    'Do FooBar
  Case Is < 10000
    'Conflict
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachable1()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is > 5000 
    'Do FooBar
  Case Is < 10000
    'Conflict
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachable2()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is > 5000 
    'Do FooBar
  Case Is <= 10000
    'Conflict
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachable3()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case z <> 5 
    'Do FooBar
  Case Is = 5
    'OK
  Case Else
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_Evil()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case 1 To 4, 7 To 9, 11, 13, Is > 400 
    'Do FooBar
  Case 401
    'Conflict
  Case 15 To 25, 300 To 401
    'Conflict
  Case z < 2
    'Conflict
  Case 88
    'OK
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 3);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_RelationalOpSimple()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstant2()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstant3()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SelectCaseUsesConstant4()
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_RelationalOpExpression()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGTFollowsRange()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGTSubsequent()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtGTE()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtLT()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtLTE()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

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
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtNEQ()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is <> 8
    'Do FooBar
  Case 7
    'Unreachable
  Case 5001
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtEQ()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is = 8
    'Do FooBar
  Case 8
    'Unreachable
'  Case 5001
    'Reachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtMultiple()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is > 8
    'Do FooBar
'  Case 8
    'OK
  Case  Is = 9
    'Unreachable
'  Case Is < 100
    'Conflict
'  Case Is < 5
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtAndRange()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is > 8
    'Do FooBar
  Case 3 To 10
    'Conflict
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtAndNegativeRange()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is < 8
    'Do FooBar
  Case -10 To -3
    'Unreachable
  Case 0
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 2);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_IsStmtAndRangeAreNegative()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case Is < -8
    'Do FooBar
  Case -10 To -3
    'Conflict
  Case 0
    'Do FooBar again
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SingleValueFollowedByIsStmt()
        {
            const string inputCode =
@"Sub Foo()

Const z As Long = 7

Select Case z
  Case 200
    'OK
  Case Is > 8
    'Conflict
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        private void CheckActualUnreachableBlockCountEqualsExpected(string inputCode, int expectedCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new UnreachableCaseInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(expectedCount, actualResults.Count());
        }
    }
}
