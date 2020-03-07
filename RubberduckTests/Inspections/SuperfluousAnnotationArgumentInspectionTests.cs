using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SuperfluousAnnotationArgumentInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [TestCase("PredeclaredId", "False")]
        [TestCase("PredeclaredId", "True")]
        [TestCase("Exposed", "False")]
        [TestCase("Exposed", "True")]
        [TestCase("ModuleDescription", @"""MyModule"", ""SomethingElse""")]
        [TestCase("Folder", @"""MyModule"", ""SomethingElse""")]
        public void TooManyArguments_ReturnsResult(string annotationName, string argumentsText)
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
        [TestCase("Exposed", "")]
        [TestCase("ModuleDescription", @"""MyModule""")]
        [TestCase("Folder", @"""MyModule""")]
        [TestCase("IgnoreTest", @"""Reasons""")]
        [TestCase("IgnoreModule", @"Inspection1, Inspection2, Inspection3, Inspection4, Inspection5, Inspection6")]
        public void NotTooManyArguments_DoesNotReturnResult(string annotationName, string argumentsText)
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
            return new SuperfluousAnnotationArgumentInspection(state);
        }
    }
}