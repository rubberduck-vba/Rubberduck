using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveStepOneQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void StepOne_QuickFixWorks_Remove()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 5
    Next
End Sub";

            TestStepOneQuickFix(expectedCode, inputCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void StepOne_QuickFixWorks_NestedLoops()
        {
            var inputCode =
@"Sub Foo()
    For value = 0 To 5
        For value = 0 To 5 Step 1
        Next
    Next
End Sub";

            var expectedCode =
@"Sub Foo()
    For value = 0 To 5
        For value = 0 To 5
        Next
    Next
End Sub";

            TestStepOneQuickFix(expectedCode, inputCode);
        }

        private void TestStepOneQuickFix(string expectedCode, string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new StepOneIsRedundantInspection(state) { Severity = CodeInspectionSeverity.Warning };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveStepOneQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }
    }
}
