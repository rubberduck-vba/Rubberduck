using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitActiveWorkbookReferenceInspectionTests
    {
        [Test]
        [Ignore("This was apparently only passing due to the test setup. See #4404")]
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
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();


            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ImplicitActiveWorkbookReferenceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                return inspectionResults.Count();
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitActiveWorkbookReferenceInspection";
            var inspection = new ImplicitActiveWorkbookReferenceInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
