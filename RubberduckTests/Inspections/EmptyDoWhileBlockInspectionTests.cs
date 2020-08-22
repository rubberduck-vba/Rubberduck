using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyDoWhileBlockInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyDoWhileBlock_InspectionName()
        {
            var inspection = new EmptyDoWhileBlockInspection(null);

            Assert.AreEqual(nameof(EmptyDoWhileBlockInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyDoWhileBlock_DoesNotFiresOnImplementedLoopBlocks()
        {
            const string inputCode =
                @"Sub Foo(results As Collection)
    Dim i As Integer
    i = 1

    Do While i < 6
        Cells( i, 1 ).Value = 20
        i = i + 1
    Loop
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyDoWhileBlock_FiresOnEmptyLoopBlocks()
        {
            const string inputCode =
                @"Sub Foo(results As Collection)
    Dim i As Integer
    i = 1

    Do While i < 6
        'Cells( i, 1 ).Value = 20
        'i = i + 1
    Loop
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyDoWhileBlockInspection(state);
        }
    }
}
