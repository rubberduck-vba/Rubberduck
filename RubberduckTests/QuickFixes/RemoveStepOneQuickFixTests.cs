using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class RemoveStepOneQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
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

            this.TestStepOneQuickFix(expectedCode, inputCode);
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
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

            this.TestStepOneQuickFix(expectedCode, inputCode);
        }

        private void TestStepOneQuickFix(string expectedCode, string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new StepOneIsRedundantInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveStepOneQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }
    }
}
