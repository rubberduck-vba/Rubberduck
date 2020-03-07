using NUnit.Framework;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class OnLocalErrorInspectionTests : InspectionTestsBase
    {
        [TestCase("On Local Error GoTo 0", 1)]
        [TestCase("On Local Error Resume Next", 1)]
        [TestCase("On Local Error GoTo Label\r\nLabel: ", 1)]
        [TestCase("On Error GoTo 0", 0)]
        [Category("Inspections")]
        public void OnLocalError_VariousScenarios(string body, int expectedCount)
        {
            string inputCode =
$@"Sub foo()
    {body}
End Sub";
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new OnLocalErrorInspection(state);
        }
    }
}
