using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class StepOneIsRedundantInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void StepOneIsRedundant_ReturnsResult()
        {
            string inputCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
    Next
End Sub";

            this.TestStepOneIsRedundantInspection(inputCode, 1);
        }

        [Test]
        [Category("Inspections")]
        public void StepOneIsRedundant_NestedLoopsAreDetected()
        {
            string inputCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
        For value = 0 To 5 Step 1
        Next
    Next
End Sub";

            this.TestStepOneIsRedundantInspection(inputCode, 2);
        }

        private void TestStepOneIsRedundantInspection(string inputCode, int expectedResultCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new StepOneIsRedundantInspection(state) { Severity = CodeInspectionSeverity.Warning };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(expectedResultCount, inspectionResults.Count());
        }
    }
}
