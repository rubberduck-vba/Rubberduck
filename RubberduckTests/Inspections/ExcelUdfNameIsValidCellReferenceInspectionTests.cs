using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ExcelUdfNameIsValidCellReferenceInspectionTests : InspectionTestsBase
    {
        [TestCase("a1")]
        [TestCase("A1")]
        [TestCase("AA1")]
        [TestCase("ZZ1")]
        [TestCase("XFD1")]
        [TestCase("XEZ1")]
        [TestCase("WZZ1")]
        [TestCase("Foo42")]
        [TestCase("XFD1048576")]        
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsResult_ValidCells(string identifier)
        {
            const string codeTemplate =
@"Public Function {0}() As Long
    {0} = 42
End Function
";

            Assert.AreEqual(1, InspectionResultCount(string.Format(codeTemplate, identifier), ComponentType.StandardModule));
        }

        [Test]
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsResult_ValidCellsInPrivateModule()
        {
            const string code =
                @"Option Private Module
Public Function A1() As Long
    A1 = 42
End Function
";

            Assert.AreEqual(0, InspectionResultCount(code, ComponentType.StandardModule));
        }

        [TestCase("Foo")]
        [TestCase("XXX69")]
        [TestCase("XKCD42")]
        [TestCase("AAA1234567")]
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsNoResult_InvalidAsCell(string identifier)
        {
            const string codeTemplate =
@"Public Function {0}() As Long
    {0} = 42
End Function
";

            Assert.AreEqual(0, InspectionResultCount(string.Format(codeTemplate, identifier), ComponentType.StandardModule));
        }

        [TestCase(ComponentType.ClassModule)]
        [TestCase(ComponentType.UserForm)]
        [TestCase(ComponentType.DocObject)]
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsNoResult_NonStandardModule(ComponentType moduleType)
        {
            const string code =
@"Public Function A1() As Long
    A1 = 42
End Function
";

            Assert.AreEqual(0, InspectionResultCount(code, moduleType));
        }

        [Test]
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsNoResult_Ignored()
        {
            const string code =
@"'@Ignore ExcelUdfNameIsValidCellReference
Public Function A1() As Long
    A1 = 42
End Function
";

            Assert.AreEqual(0, InspectionResultCount(code, ComponentType.StandardModule));
        }

        [Test]
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsNoResult_PrivateFunction()
        {
            const string code =
@"Private Function A1() As Long
    A1 = 42
End Function
";

            Assert.AreEqual(0, InspectionResultCount(code, ComponentType.StandardModule));
        }

        [TestCase("Sub A1()", "Sub")]
        [TestCase("Property Get A1() As Long", "Property")]
        [TestCase("Property Let A1(foo As Long)", "Property")]
        [TestCase("Property Set A1(foo As Variant)", "Property")]
        [Category("Inspections")]
        public void ExcelUdfNameIsValidCellReferenceInspection_ReturnsNoResult_NonFunction(string signature, string ending)
        {
            const string codeTemplate =
@"{0}
    A1 = 42
End {1}
";

            Assert.AreEqual(0, InspectionResultCount(string.Format(codeTemplate, signature, ending), ComponentType.StandardModule));
        }

        private int InspectionResultCount(string inputCode, ComponentType moduleType)
            => InspectionResultsForModules(("UnderTest", inputCode, moduleType), ReferenceLibrary.Excel).Count();

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ExcelUdfNameIsValidCellReferenceInspection(state);
        }
    }
}
