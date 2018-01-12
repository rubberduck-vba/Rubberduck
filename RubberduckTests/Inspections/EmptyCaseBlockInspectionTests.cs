using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyCaseBlockInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void EmptyCaseBlock_InspectionType()
        {
            var inspection = new EmptyCaseBlockInspection(null);
            var expectedInspection = CodeInspectionType.MaintainabilityAndReadabilityIssues;

            Assert.AreEqual(expectedInspection, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyCaseBlock_InspectionName()
        {
            const string expectedName = nameof(EmptyCaseBlockInspection);
            var inspection = new EmptyCaseBlockInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyCaseBlock_DoesNotFiresOnImplementedCaseBlocks()
        {
            const string inputCode =
                @"Sub Foo(caseNum As Long)
    Select Case caseNum
        Case 1
            MsgBox ""1""
        Case 2
            MsgBox ""2""
        Case 3
            MsgBox ""3""
        Case Else
            MsgBox ""?""
    End Select
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyCaseBlock_FiresOnEmptyCaseBlocks()
        {
            const string inputCode =
                @"Sub Foo(caseNum As Long)
    Select Case caseNum
        Case 1
        Case 2
        Case 3
        Case Else
            MsgBox ""?""
    End Select
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 3);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyCaseBlock_FiresOnCommentCaseBlocks()
        {
            const string inputCode =
                @"Sub Foo(caseNum As Long)
    Select Case caseNum
        Case 1
            'TODO - handle this case!
        Case Else
            MsgBox ""?""
    End Select
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 1);
        }

        private void CheckActualEmptyBlockCountEqualsExpected(string inputCode, int expectedCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyCaseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }
    }
}
