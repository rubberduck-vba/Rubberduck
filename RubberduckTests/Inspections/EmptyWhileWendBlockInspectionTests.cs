using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyWhileWendBlockInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyWhileWendBlock_InspectionName()
        {
            var inspection = new EmptyWhileWendBlockInspection(null);

            Assert.AreEqual(nameof(EmptyWhileWendBlockInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
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
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyWhileWendBlockInspection(state);
        }
    }
}
