using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ModuleWithoutFolderInspectionTests : InspectionTestsBase
    {
        [TestCase("Option Explicit", 1)] //NoFolderAnnotation
        [TestCase("'@Folder Foo.Bar\r\nOption Explicit", 0)] //FolderAnnotation
        [TestCase("'@PredeclaredId\r\nOption Explicit", 1)] //NonFolderAnnotation
        [TestCase("'@IgnoreModule ModuleWithoutFolder\r\nOption Explicit", 0)] //NoFolderAnnotation_IgnoreWorks
        [Category("Inspections")]
        public void Module_VariousScenarios(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ModuleWithoutFolderInspection(state);
        }
    }
}
