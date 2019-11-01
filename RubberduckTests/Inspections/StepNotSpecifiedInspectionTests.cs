using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class StepNotSpecifiedInspectionTests : InspectionTestsBase
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

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
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

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new StepIsNotSpecifiedInspection(state) { Severity = CodeInspectionSeverity.Warning };
        }
    }
}
