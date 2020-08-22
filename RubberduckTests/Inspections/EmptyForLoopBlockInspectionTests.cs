using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyForLoopBlockInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyForLoopBlock_InspectionName()
        {
            var inspection = new EmptyForLoopBlockInspection(null);

            Assert.AreEqual(nameof(EmptyForLoopBlockInspection), inspection.Name);
        }

        [TestCase("idx = idx + 2", 0)]
        [TestCase("", 1)]
        [Category("Inspections")]
        public void EmptyForLoopBlock_LoopBlockContentScenarios(string forLoopContent, int expectedCount)
        {
            string inputCode =
                $@"Sub Foo()
    Dim idx As Integer
    for idx = 1 to 100
        {forLoopContent}
    next idx
End Sub";
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyForLoopBlockInspection(state);
        }
    }
}
