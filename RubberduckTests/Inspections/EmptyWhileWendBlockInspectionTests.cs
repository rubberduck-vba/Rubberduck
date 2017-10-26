using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;

namespace RubberduckTests.Inspections
{
    [TestClass, Ignore]
    public class EmptyWhileWendBlockInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyWhileWendBlock_InspectionType()
        {
            var inspection = new EmptyWhileWendBlockInspection(null);
            var expectedInspection = CodeInspectionType.MaintainabilityAndReadabilityIssues;

            Assert.AreEqual(expectedInspection, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyWhileWendBlock_InspectionName()
        {
            const string expectedName = nameof(EmptyWhileWendBlockInspection);
            var inspection = new EmptyWhileWendBlockInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyWhileWendBlock_DoesNotFiresOnImplementedLoopBlocks()
        {
            const string inputCode =
                @"Sub Foo(results As Collection)
    Dim LTotal As Integer

    LTotal = 1

    While LTotal < 5
        MsgBox (LTotal)
        LTotal = LTotal + 1
    Wend
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 0);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyWhileWendBlock_FiresOnEmptyLoopBlocks()
        {
            const string inputCode =
                @"Sub Foo(results As Collection)
    Dim LTotal As Integer

    LTotal = 1

    While LTotal < 5
        'MsgBox (LTotal)
        'LTotal = LTotal + 1
    Wend
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 1);
        }

        private void CheckActualEmptyBlockCountEqualsExpected(string inputCode, int expectedCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyWhileWendBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }
    }
}
