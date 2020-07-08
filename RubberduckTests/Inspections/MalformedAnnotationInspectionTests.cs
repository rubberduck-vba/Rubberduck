using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MalformedAnnotationInspectionTests : InspectionTestsBase
    {
        [TestCase(@"'@Folder", 1)]
        [TestCase(@"'@Folder ""Foo""", 0)]
        [TestCase(@"'@Ignore", 1)]
        [TestCase(@"'@Ignore ProcedureNotUsedInspection", 0)]
        [TestCase("'@Folder\r\n'@Ignore", 2)]
        [Category("Inspections")]
        public void MalformedAnnotation_Various(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new MissingAnnotationArgumentInspection(null);

            Assert.AreEqual(nameof(MissingAnnotationArgumentInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MissingAnnotationArgumentInspection(state);
        }
    }
}
