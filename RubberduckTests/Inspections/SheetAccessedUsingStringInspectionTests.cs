using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SheetAccessedUsingStringInspectionTests
    {
        [Test]
        [Ignore("See #4411")]
        [Category("Inspections")]
        public void SheetAccessedUsingString_ReturnsResult_AccessingUsingWorkbookModule()
        {
            const string inputCode =
                @"Public Sub Foo()
    ThisWorkbook.Worksheets(""Sheet1"").Range(""A1"") = ""Foo""
    ThisWorkbook.Sheets(""Sheet1"").Range(""A1"") = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void SheetAccessedUsingString_ReturnsResult_AccessingUsingApplicationModule()
        {
            const string inputCode =
                @"Public Sub Foo()
    Application.Worksheets(""Sheet1"").Range(""A1"") = ""Foo""
    Application.Sheets(""Sheet1"").Range(""A1"") = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void SheetAccessedUsingString_ReturnsResult_AccessingUsingGlobalModule()
        {
            const string inputCode =
                @"Public Sub Foo()
    Worksheets(""Sheet1"").Range(""A1"") = ""Foo""
    Sheets(""Sheet1"").Range(""A1"") = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Ignore("Ref #4329")]
        [Category("Inspections")]
        public void SheetAccessedUsingString_DoesNotReturnResult_AccessingUsingWorkbooksProperty()
        {
            const string inputCode =
                @"Public Sub Foo()
    Workbooks(""Foo"").Worksheets(""Sheet1"").Range(""A1"") = ""Foo""
    Workbooks(""Foo"").Sheets(""Sheet1"").Range(""A1"") = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void SheetAccessedUsingString_DoesNotReturnResult_NoSheetWithGivenNameExists()
        {
            const string inputCode =
                @"Public Sub Foo()
    ThisWorkbook.Worksheets(""BadName"").Range(""A1"") = ""Foo""
    ThisWorkbook.Sheets(""BadName"").Range(""A1"") = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void SheetAccessedUsingString_DoesNotReturnResult_SheetWithGivenNameExistsInAnotherProject()
        {
            const string inputCode =
                @"Public Sub Foo()
    ThisWorkbook.Worksheets(""SheetFromOtherProject"").Range(""A1"") = ""Foo""
    ThisWorkbook.Sheets(""SheetFromOtherProject"").Range(""A1"") = ""Foo""
End Sub";

            // Referenced project is created inside helper method
            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void SheetAccessedUsingString_DoesNotReturnResult_AccessingUsingVariable()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim s As String
    s = ""Sheet1""

    ThisWorkbook.Worksheets(s).Range(""A1"") = ""Foo""
    ThisWorkbook.Sheets(s).Range(""A1"") = ""Foo""
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new SheetAccessedUsingStringInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        private static RubberduckParserState ArrangeParserAndParse(string inputCode)
        {
            var builder = new MockVbeBuilder();

            var referencedProject = builder.ProjectBuilder("ReferencedProject", ProjectProtection.Unprotected)
                .AddComponent("SheetFromOtherProject", ComponentType.Document, "",
                    properties: new[]
                    {
                        CreateVBComponentPropertyMock("Name", "SheetFromOtherProject").Object,
                        CreateVBComponentPropertyMock("CodeName", "SheetFromOtherProject").Object
                    })
                .Build();

            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddComponent("Sheet1", ComponentType.Document, "",
                    properties: new[]
                    {
                        CreateVBComponentPropertyMock("Name", "Sheet1").Object,
                        CreateVBComponentPropertyMock("CodeName", "Sheet1").Object
                    })
                .AddReference("ReferencedProject", string.Empty, 0, 0)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(referencedProject).AddProject(project).Build();

            return MockParser.CreateAndParse(vbe.Object);
        }

        // ReSharper disable once InconsistentNaming
        private static Mock<IProperty> CreateVBComponentPropertyMock(string propertyName, string propertyValue)
        {
            var propertyMock = new Mock<IProperty>();
            propertyMock.SetupGet(m => m.Name).Returns(propertyName);
            propertyMock.SetupGet(m => m.Value).Returns(propertyValue);

            return propertyMock;
        }
    }
}
