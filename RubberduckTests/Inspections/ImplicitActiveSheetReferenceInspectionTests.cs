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
    public class ImplicitActiveSheetReferenceInspectionTests : InspectionTestsBase
    {
        private static readonly IDictionary<string, IEnumerable<string>> DefaultDocumentModuleSupertypeNames = new Dictionary<string, IEnumerable<string>>
        {
            ["ThisWorkbook"] = new[] { "Workbook", "_Workbook" },
            ["Sheet1"] = new[] { "Worksheet", "_Worksheet" }
        };

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
            var modules = new(string, string, ComponentType)[] 
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Class1", inputCode, ComponentType.ClassModule)
            };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Class1", inputCode, ComponentType.ClassModule)
            };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Class1", inputCode, ComponentType.ClassModule)
            };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Class1", inputCode, ComponentType.ClassModule)
            };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", inputCode, ComponentType.Document),
                ("Class1", string.Empty, ComponentType.ClassModule)
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", inputCode, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Class1", string.Empty, ComponentType.ClassModule)
            };
            Assert.AreEqual(1, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_NoResultForWorksheetVariable()
        {
            const string inputCode =
                @"Sub foo()
    Dim sh As Worksheet
    Set sh = Sheet1
    Debug.Print sh.Cells(1, 1)
End Sub
";
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Module1", inputCode, ComponentType.StandardModule)
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_NoResultForWorksheetFunction()
        {
            const string inputCode =
                @"Sub foo()
    Debug.Print GetSheet.Cells(1, 1)
End Sub

Private Function GetSheet() As Worksheet
End Function
";
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Module1", inputCode, ComponentType.StandardModule)
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitActiveSheetReference_NoResultForWorksheetProperty()
        {
            const string inputCode =
                @"Sub foo()
    Debug.Print GetSheet.Cells(1, 1)
End Sub

Private Property Get GetSheet() As Worksheet
End Property
";
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Module1", inputCode, ComponentType.StandardModule)
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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
            var modules = new (string, string, ComponentType)[]
            {
                ("ThisWorkbook", string.Empty, ComponentType.Document),
                ("Sheet1", string.Empty, ComponentType.Document),
                ("Class1", inputCode, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.Excel, DefaultDocumentModuleSupertypeNames).Count());
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