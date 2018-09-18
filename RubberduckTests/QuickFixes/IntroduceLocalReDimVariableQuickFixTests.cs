using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IntroduceLocalReDimVariableQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_OnlyFixesFixesFirstOccurrenceOfReDimStatement()
        {
            var inputCode = @"
Public Sub DoSomething()
    ReDim foo(1)
    ReDim foo(2)
End Sub
";
            var expectedCode = @"
Public Sub DoSomething()
    Dim foo As Variant
    ReDim foo(1)
    ReDim foo(2)
End Sub
";
            TestInsertLocalReDimVariableQuickFix(expectedCode, inputCode);
        }

        [Ignore("Type inference isn't implemented. Un-ignore this test when it is.")]
        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_KeepsAsClauseIfConsistent()
        {
            var inputCode = @"
Public Sub DoSomething()
    ReDim foo(1) As Long
    ReDim foo(2) As Long
End Sub
";
            var expectedCode = @"
Public Sub DoSomething()
    Dim foo As Long
    ReDim foo(1) As Long
    ReDim foo(2) As Long
End Sub
";
            TestInsertLocalReDimVariableQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceLocalVariable_DeclaresAsVariantIfInconsistentAsClauses()
        {
            var inputCode = @"
Public Sub DoSomething()
    ReDim foo(1) As Long
    ReDim foo(2) As String
End Sub
";
            var expectedCode = @"
Public Sub DoSomething()
    Dim foo As Variant
    ReDim foo(1) As Long
    ReDim foo(2) As String
End Sub
";
            TestInsertLocalReDimVariableQuickFix(expectedCode, inputCode);
        }
        private void TestInsertLocalReDimVariableQuickFix(string expectedCode, string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new UndeclaredRedimVariableInspection(state) { Severity = CodeInspectionSeverity.Warning };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IntroduceLocalVariableQuickFix(state).Fix(inspectionResults.First());
            var actualCode = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actualCode);
        }
    }
}