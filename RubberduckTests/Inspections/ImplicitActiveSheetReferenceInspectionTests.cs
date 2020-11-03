using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

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
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_ReportsCells()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Cells(1,2)
End Sub
";
            var modules = new (string, string, ComponentType)[] { ("Class1", inputCode, ComponentType.ClassModule) };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_ReportsColumns()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Columns(3)
End Sub
";
            var modules = new (string, string, ComponentType)[] { ("Class1", inputCode, ComponentType.ClassModule) };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_ReportsRows()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Rows(3)
End Sub
";
            var modules = new (string, string, ComponentType)[] { ("Class1", inputCode, ComponentType.ClassModule) };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_DoesNotReportInWorkSheetModules()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Cells(1,2)
End Sub
";
            var module = ("Sheet1", inputCode, ComponentType.Document);
            var vbe = MockVbeBuilder.BuildFromModules(module, ReferenceLibrary.Excel).Object;

            using (var state = MockParser.CreateAndParse(vbe))
            {
                var documentModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Document)
                    .OfType<DocumentModuleDeclaration>()
                    .Single();
                documentModule.AddSupertypeName("Worksheet");

                var inspection = InspectionUnderTest(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_ReportsInWorkbookModules()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Cells(1,2)
End Sub
";
            var module = ("Sheet1", inputCode, ComponentType.Document);
            var vbe = MockVbeBuilder.BuildFromModules(module, ReferenceLibrary.Excel).Object;

            using (var state = MockParser.CreateAndParse(vbe))
            {
                var documentModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Document)
                    .OfType<DocumentModuleDeclaration>()
                    .Single();
                documentModule.AddSupertypeName("Workbook");

                var inspection = InspectionUnderTest(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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
            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count());
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