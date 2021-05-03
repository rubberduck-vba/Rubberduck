using System.Collections.Generic;
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
    public class ImplicitContainingWorkbookReferenceInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_ReportsWorksheets()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";
            Assert.AreEqual(1, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_ExplicitApplication()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Application.Worksheets(""Sheet1"")
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
            ;
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_ReportsSheets()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Sheets(""Sheet1"")
End Sub";
            Assert.AreEqual(1, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_ReportsNames()
        {
            const string inputCode =
                @"
Sub foo()
    Names.Add ""foo"", Rows(1)
End Sub";
            Assert.AreEqual(1, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_ExplicitReference_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim book As Workbook
    Dim sheet As Worksheet
    Set sheet = book.Worksheets(1)
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_MeReference_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Me.Worksheets(1)
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_WithReference_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    With Me
        Set sheet = .Worksheets(1)
    End With
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_DimAsTypeWorksheets_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim allSheets As Worksheets
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_DimAsTypeSheets_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim allSheets As Sheets
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_DimAsTypeNames_NotReported()
        {
            const string inputCode =
                @"
Sub foo()
    Dim allNames As Names
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_DoesNotReportUnqualifiedOutsideWorkbookModules()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";
            const int expected = 0;
            var actual = ArrangeOutsideWorkbookAndGetInspectionCount(inputCode);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingWorkbookReference_Ignored_DoesNotReportRange()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet

    '@Ignore ImplicitContainingWorkbookReference
    Set sheet = Worksheets(""Sheet1"")
End Sub";
            Assert.AreEqual(0, InspectionResultsInWorkbook(inputCode).Count());
        }

        private int ArrangeOutsideWorkbookAndGetInspectionCount(string code)
        {
            var modules = new (string, string, ComponentType)[] {("Module1", code, ComponentType.StandardModule)};
            return InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count();
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ImplicitContainingWorkbookReferenceInspection(null);

            Assert.AreEqual(nameof(ImplicitContainingWorkbookReferenceInspection), inspection.Name);
        }

        private IEnumerable<IInspectionResult> InspectionResultsInWorkbook(string inputCode)
        {
            // a VBA project hosted in Excel always has a ThisWorkbook module and AT LEAST one Worksheet module (default: "Sheet1").
            var defaultDocumentModuleSupertypeNames = new Dictionary<string, IEnumerable<string>>
            {
                ["ThisWorkbook"] = new[] { "Workbook", "_Workbook" },
                ["Sheet1"] = new[] { "Worksheet", "_Worksheet" }
            };

            var modules = new[] {
                ("Sheet1", string.Empty, ComponentType.Document),
                ("ThisWorkbook", inputCode, ComponentType.Document)
            };
            var vbe = MockVbeBuilder.BuildFromModules(modules, new[] { ReferenceLibrary.Excel }).Object;

            using(var state = MockParser.CreateAndParse(vbe, documentModuleSupertypeNames:defaultDocumentModuleSupertypeNames))
            {
                var inspection = InspectionUnderTest(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitContainingWorkbookReferenceInspection(state);
        }
    }
}