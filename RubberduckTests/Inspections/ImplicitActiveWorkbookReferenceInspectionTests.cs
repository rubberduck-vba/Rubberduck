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
    public class ImplicitActiveWorkbookReferenceInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ReportsWorksheets()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ExplicitApplication()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Application.Worksheets(""Sheet1"")
End Sub";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ReportsSheets()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Sheets(""Sheet1"")
End Sub";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ReportsNames()
        {
            const string inputCode =
                @"
Sub foo()
    Names.Add ""foo"", Rows(1)
End Sub";

            const int expected = 1;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ExplicitReference_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim book As Workbook
    Dim sheet As Worksheet
    Set sheet = book.Worksheets(1)
End Sub";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ExplicitParameterReference_NotReported()
        {
            const string inputCode =
                @"
Sub foo(book As Workbook)
    Debug.Print book.Worksheets.Count
End Sub";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_DimAsTypeWorksheets_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim allSheets As Worksheets
End Sub";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_DimAsTypeSheets_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim allSheets As Sheets
End Sub";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_DimAsTypeNames_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim allNames As Names
End Sub";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_DoesNotReportUnqualifiedInWorkbookModules()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";
            var module = ("SomeWorkbook", inputCode, ComponentType.Document);
            var vbe = MockVbeBuilder.BuildFromModules(module, ReferenceLibrary.Excel).Object;

            using (var state = MockParser.CreateAndParse(vbe))
            {
                var documentModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Document)
                    .OfType<DocumentModuleDeclaration>()
                    .Single();
                documentModule.AddSupertypeName("Workbook");

                var inspection = InspectionUnderTest(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_ReportsApplicationQualifiedInWorkbookModules()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Application.Worksheets(""Sheet1"")
End Sub";
            var module = ("SomeWorkbook", inputCode, ComponentType.Document);
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
        public void ImplicitActiveWorkbookReference_ReportsInWorksheetModules()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";
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

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveWorkbookReference_Ignored_DoesNotReportRange()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet

    '@Ignore ImplicitActiveWorkbookReference
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            const int expected = 0;
            var actual = ArrangeAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        private int ArrangeAndGetInspectionCount(string code)
        {
            var modules = new(string, string, ComponentType)[] { ("Module1", code, ComponentType.StandardModule) };
            return InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count();
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ImplicitActiveWorkbookReferenceInspection(null);

            Assert.AreEqual(nameof(ImplicitActiveWorkbookReferenceInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitActiveWorkbookReferenceInspection(state);
        }
    }
}
