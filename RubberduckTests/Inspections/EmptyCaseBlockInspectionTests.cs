using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyCaseBlockInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyCaseBlock_InspectionName()
        {
            var inspection = new EmptyCaseBlockInspection(null);

            Assert.AreEqual(nameof(EmptyCaseBlockInspection), inspection.Name);
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
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
            Assert.AreEqual(3, InspectionResultsForStandardModule(inputCode).Count());
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
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyCaseBlockInspection(state);
        }
    }
}
