using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ModuleScopeDimKeywordInspectionTests : InspectionTestsBase
    {
        [TestCase("Dim foo As String", 1)]
        [TestCase("Dim foo\r\nDim bar", 2)]
        [TestCase("Private foo", 0)]
        [TestCase("'@IgnoreModule\r\nDim foo", 0)]
        [TestCase("'@IgnoreModule ModuleScopeDimKeyword\r\nDim foo", 0)]
        [TestCase("'@IgnoreModule VariableNotUsed\r\nDim foo", 1)]
        [TestCase("'@IgnoreModule ModuleScopeDimKeyword\r\nDim foo", 0)]
        [Category("Inspections")]
        public void ModuleScopeDimKeyword_ReturnsResult(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ModuleScopeDimKeywordInspection(null);

            Assert.AreEqual(nameof(ModuleScopeDimKeywordInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ModuleScopeDimKeywordInspection(state);
        }
    }
}
