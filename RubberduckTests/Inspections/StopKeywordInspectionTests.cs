using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class StopKeywordInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void StopKeyword_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Stop
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void NoStopKeyword_NoResult()
        {
            var inputCode =
                @"Sub Foo()
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void StopKeyword_Ignored_DoesNotReturnResult()
        {
            var inputCode =
                @"Sub Foo()
'@Ignore StopKeyword
    Stop
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void StopKeywords_Ignored_ReturnsCorrectResults()
        {
            var inputCode =
                @"Sub Foo()
    Dim d As Integer
    d = 0
    Stop

    d = 1

    '@Ignore StopKeyword
    Stop
End Sub";
            var results = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, results.Count());
            Assert.AreEqual(4, results.First().QualifiedSelection.Selection.StartLine);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new StopKeywordInspection(null);

            Assert.AreEqual(nameof(StopKeywordInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new StopKeywordInspection(state);
        }
    }
}
