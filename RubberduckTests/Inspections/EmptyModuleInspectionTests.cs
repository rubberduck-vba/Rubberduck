using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyModuleInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ModuleWithContentNotRepresentingFunctionality_ReturnsResult()
        {
            const string inputCode =
                @"Option Base 1
Option Compare Binary
Option Explicit
Option Private Module

'Nothing to see here. 

DefBool B: DefByte Y: DefInt I: DefLng L: DefLngLng N: DefLngPtr P: DefCur C: DefSng G: DefDbl D: DefDate T: DefStr E: DefObj O: DefVar V

'Here, neither! _
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ClassWithContentNotRepresentingFunctionality_ReturnsResult()
        {
            const string inputCode =
                @"Option Base 1
Option Compare Binary
Option Explicit

'Nothing to see here. 

DefBool B: DefByte Y: DefInt I: DefLng L: DefLngLng N: DefLngPtr P: DefCur C: DefSng G: DefDbl D: DefDate T: DefStr E: DefObj O: DefVar V

'Here, neither! _
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase("Private Function Foo() As String\r\nEnd Function")]
        [TestCase("Private Sub Foo()\r\nEnd Sub")]
        [TestCase("Public Property Get Foo()\r\nEnd Property")]
        [TestCase("Public Property Set Foo(rhs As Variant)\r\nEnd Property")]
        [TestCase("Public Property Let Foo(rhs As Variant)\r\nEnd Property")]
        [TestCase("Private Enum Foo\r\nBar\r\nEnd Enum")]
        [TestCase("Private Type Foo\r\nBar As String\r\nEnd Type")]
        [TestCase("Public Type Foo\r\nBar As String\r\nEnd Type")]
        [TestCase("Private foo As String")]
        [TestCase("Private Const foo As Long = 6")]
        [TestCase("Public Event Foo(bar As Variant)")]
        [TestCase("'@IgnoreModule EmptyModule")]
        [Category("Inspections")]
        public void ModulesWithVariousContent_NoResults(string inputCode)
        {
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase("")]
        [TestCase("'@IgnoreModule EmptyModule")]
        [Category("Inspections")]
        public void EmptyDocumentModules_NoResults(string inputCode)
        {
            Assert.AreEqual(0, InspectionResultsForModules(("TestDoc", inputCode, ComponentType.Document)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new EmptyModuleInspection(null, null);

            Assert.AreEqual(nameof(EmptyModuleInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyModuleInspection(state, state);
        }
    }
}
