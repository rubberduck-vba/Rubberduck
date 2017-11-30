using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class StepNotSpecifiedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void StepNotSpecified_ReturnsResult()
        {
            string inputCode =
@"Sub Foo()
    For value = 0 To 5
    Next
End Sub";

            this.TestStepNotSpecifiedInspection(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void StepNotSpecified_NestedLoopsAreDetected()
        {
            string inputCode =
@"Sub Foo()
    For value = 0 To 5
        For value = 0 To 5
        Next
    Next
End Sub";

            this.TestStepNotSpecifiedInspection(inputCode, 2);
        }

        private void TestStepNotSpecifiedInspection(string inputCode, int expectedResultCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new StepIsNotSpecifiedInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(expectedResultCount, inspectionResults.Count());
        }
    }
}
