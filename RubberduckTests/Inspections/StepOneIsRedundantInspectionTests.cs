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
    public class StepOneIsRedundantInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void StepOneIsRedundant_ReturnsResult()
        {
            string inputCode =
@"Sub Foo()
    For value = 0 To 5 Step 1
    Next
End Sub";

            this.TestStepOneIsRedundantInspection(inputCode, 1);
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var inspection = new StepOneIsRedundantInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(expectedResultCount, inspectionResults.Count());
        }
    }
}
