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
  Case 1 to 100
    'Do FooBar
  Case 50
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_SimpleLongCollision()
        {
            const string inputCode =
@"Sub Foo()

Const x as Long = 7

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
@"Sub Foo()

Const x as Long = 7

Select Case x
  Case 1 to 49
    'Do FooBar
  Case 50
    'Reachable
  Case ""Test""
    'Unreachable
End Select

End Sub";
            CheckActualUnreachableBlockCountEqualsExpected(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void UnreachableCaseInspection_MultipleRanges()
        {
            const string inputCode =
@"Sub Foo()

Const x As Long = 7
Const maxNumber As Long = 5000

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
        public void UnreachableCaseInspection_IsStmt()
        {
            const string inputCode =
@"Sub Foo()

Const x As Long = 7
'Const maxNumber As Long = 5000

Select Case x
  Case Is > 5000
    'Do FooBar
  Case 7
    'Do Foobar again
  Case 10000
    'Unreachable
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
