using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MissingAnnotationArgumentInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [TestCase("ModuleDescription", "")]
        [TestCase("ModuleAttribute", "AttributeName")]
        [TestCase("Folder", "")]
        [TestCase("Ignore", "")]
        public void TooFewArguments_ReturnsResult(string annotationName, string argumentsText)
        {
            var class1 =
                $@"

'@{annotationName} {argumentsText}

";

            var inspectionResults = InspectionResultsForModules(("Class1", class1, ComponentType.ClassModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("PredeclaredId", "")]
        [TestCase("ModuleDescription", @"""MyModule""")]
        [TestCase("ModuleAttribute", @"AttributeName, 42")]
        [TestCase("Folder", @"""MyModule""")]
        [TestCase("IgnoreModule", @"Inspection1")]
        [TestCase("IgnoreModule", "")]
        public void NotTooFewArguments_DoesNotReturnResult(string annotationName, string argumentsText)
        {
            var class1 =
                $@"

'@{annotationName} {argumentsText}

";

            var inspectionResults = InspectionResultsForModules(("Class1", class1, ComponentType.ClassModule));

            Assert.AreEqual(0, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MissingAnnotationArgumentInspection(state);
        }

    }
}