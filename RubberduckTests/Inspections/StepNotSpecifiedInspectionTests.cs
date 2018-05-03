using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class StepNotSpecifiedInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void StepNotSpecified_ReturnsResult()
        {
            string inputCode =
@"Sub Foo()
    For value = 0 To 5
    Next
End Sub";

            this.TestStepNotSpecifiedInspection(inputCode, 1);
        }

        [Test]
        [Category("Inspections")]
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

            var inspection = new StepIsNotSpecifiedInspection(state) { Severity = CodeInspectionSeverity.Warning };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(expectedResultCount, inspectionResults.Count());
        }
    }
}
