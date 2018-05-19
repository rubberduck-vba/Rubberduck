using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;
using System.Threading;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class AddStepOneQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void AddStepOne_QuickFixWorks_Remove()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 5
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
    Next
End Sub";

            this.TestAddStepOneQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AddStepOne_QuickFixWorks_NestedLoops()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 5
        For value = 0 To 5
        Next
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
        For value = 0 To 5 Step 1
        Next
    Next
End Sub";

            this.TestAddStepOneQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AddStepOne_QuickFixWorks_ComplexExpression()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 1 + 2
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 1 + 2 Step 1
    Next
End Sub";

            this.TestAddStepOneQuickFix(expectedCode, inputCode);
        }

        private void TestAddStepOneQuickFix(string expectedCode, string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new StepIsNotSpecifiedInspection(state) { Severity = CodeInspectionSeverity.Warning };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            foreach (var inspectionResult in inspectionResults)
            {
                new AddStepOneQuickFix(state).Fix(inspectionResult);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }
    }
}
