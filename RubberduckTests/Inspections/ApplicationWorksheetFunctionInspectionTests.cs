using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ApplicationWorksheetFunctionInspectionTests : InspectionTestsBase
    {
        private IEnumerable<IInspectionResult> GetInspectionResultsUsingExcelLibrary(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build().Object;
            return InspectionResults(vbe);
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_ReturnsResult_GlobalApplication()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub
";
            Assert.AreEqual(1, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_ReturnsResult_WithGlobalApplication()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    With Application
        foo = .Pi
    End With
End Sub
";
            Assert.AreEqual(1, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_ReturnsResult_ApplicationVariable()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    Dim xlApp as Excel.Application
    Set xlApp = Application
    foo = xlApp.Pi
End Sub
";
            Assert.AreEqual(1, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_ReturnsResult_WithApplicationVariable()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    Dim xlApp as Excel.Application
    Set xlApp = Application
    With xlApp
        foo = .Pi
    End With
End Sub
";
            Assert.AreEqual(1, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_DoesNotReturnResult_ExplicitUseGlobalApplication()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    foo = Application.WorksheetFunction.Pi
End Sub
";
            Assert.AreEqual(0, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_DoesNotReturnResult_ExplicitUseApplicationVariable()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    Dim xlApp as Excel.Application
    Set xlApp = Application
    foo = xlApp.WorksheetFunction.Pi
End Sub
";
            Assert.AreEqual(0, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_DoesNotReturnResult_NoExcelReference()
        {
            const string inputCode =
                @"Sub NonExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    '@Ignore ApplicationWorksheetFunction
    foo = Application.Pi
End Sub
";
            Assert.AreEqual(0, GetInspectionResultsUsingExcelLibrary(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ApplicationWorksheetFunctionInspection";
            var inspection = new ApplicationWorksheetFunctionInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ApplicationWorksheetFunctionInspection(state);
        }
    }
}