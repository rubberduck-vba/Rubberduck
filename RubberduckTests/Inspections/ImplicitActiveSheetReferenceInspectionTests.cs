using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitActiveSheetReferenceInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_ReportsRange()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant
    arr1 = Range(""A1:B2"")
End Sub
";
            var modules = new(string, string, ComponentType)[] { ("Class1", inputCode, ComponentType.ClassModule) };
            Assert.AreEqual(1, InspectionResultsForModules(modules, "Excel").Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_Ignored_DoesNotReportRange()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant

    '@Ignore ImplicitActiveSheetReference
    arr1 = Range(""A1:B2"")
End Sub
";

            var modules = new(string, string, ComponentType)[] { ("Class1", inputCode, ComponentType.ClassModule) };
            Assert.AreEqual(0, InspectionResultsForModules(modules, "Excel").Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ImplicitActiveSheetReferenceInspection(null);

            Assert.AreEqual(nameof(ImplicitActiveSheetReferenceInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitActiveSheetReferenceInspection(state);
        }
    }
}