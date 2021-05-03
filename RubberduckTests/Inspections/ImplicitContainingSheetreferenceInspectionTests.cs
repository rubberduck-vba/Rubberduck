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
    public class ImplicitContainingSheetReferenceInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_ReportsRangeInWorksheets()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant
    arr1 = Range(""A1:B2"")
End Sub
";
            Assert.AreEqual(1, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_ReportsCellsInWorksheets()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Cells(1,2)
End Sub
";
            Assert.AreEqual(1, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_ReportsColumnsInWorksheets()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Columns(3)
End Sub
";
            Assert.AreEqual(1, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_ReportsRowsInWorksheets()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Rows(3)
End Sub
";
            Assert.AreEqual(1, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_DoesNotReportsMembersQualifiedWithMe()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Me.Rows(3)
End Sub
";
            Assert.AreEqual(0, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_DoesNotReportOutsideWorkSheetModules()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Cells(1,2)
End Sub
";
            var modules = new (string, string, ComponentType)[] {
                ("Class1", inputCode, ComponentType.ClassModule),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("ThisWorkbook", string.Empty, ComponentType.Document)
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_Ignored_DoesNotReportRange()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1() As Variant

    '@Ignore ImplicitContainingWorksheetReference
    arr1 = Range(""A1:B2"")
End Sub
";
            Assert.AreEqual(0, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_NoResultForWorksheetVariable()
        {
            const string inputCode =
@"Sub foo()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(""Sheet1"")
    arr1 = sh.Range(""A1:B2"")
End Sub
";
            Assert.AreEqual(0, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_NoResultForWorksheetFunction()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1 As Variant
    arr1 = GetSheet.Range(""A1:B2"")
End Sub

Function GetSheet() As Worksheet
End Function
";
            Assert.AreEqual(0, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitContainingSheetReference_NoResultForWorksheetProperty()
        {
            const string inputCode =
@"Sub foo()
    Dim arr1 As Variant
    arr1 = GetSheet.Range(""A1:B2"")
End Sub

Property Get GetSheet() As Worksheet
End Property
";
            Assert.AreEqual(0, InspectionResultsInWorksheet(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ImplicitContainingWorksheetReferenceInspection(null);

            Assert.AreEqual(nameof(ImplicitContainingWorksheetReferenceInspection), inspection.Name);
        }

        private IEnumerable<IInspectionResult> InspectionResultsInWorksheet(string inputCode)
        {
            // a VBA project hosted in Excel always has a ThisWorkbook module and AT LEAST one Worksheet module (default: "Sheet1").
            var defaultDocumentModuleSupertypeNames = new Dictionary<string, IEnumerable<string>>
            {
                ["ThisWorkbook"] = new[] { "Workbook", "_Workbook" },
                ["Sheet1"] = new[] { "Worksheet", "_Worksheet" }
            };

            var modules = new[] {
                ("Sheet1", inputCode, ComponentType.Document),
                ("ThisWorkbook", string.Empty, ComponentType.Document)
            };
            var vbe = MockVbeBuilder.BuildFromModules(modules, new[] { ReferenceLibrary.Excel }).Object;

            using (var state = MockParser.CreateAndParse(vbe, documentModuleSupertypeNames:defaultDocumentModuleSupertypeNames))
            {
                var inspection = InspectionUnderTest(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitContainingWorksheetReferenceInspection(state);
        }
    }
}