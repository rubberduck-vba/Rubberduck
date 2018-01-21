using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Common;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ApplicationWorksheetFunctionInspectionTests
    {
        private static RubberduckParserState ArrangeParserAndParse(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object);

            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            return parser.State;
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_ReturnsResult_GlobalApplication()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub
";

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
        [Category("Inspections")]
        public void ApplicationWorksheetFunction_DoesNotReturnResult_ExplicitUseGlobalApplication()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    foo = Application.WorksheetFunction.Pi
End Sub
";

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
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
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .Build();

            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [DeploymentItem(@"Testfiles\")]
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

            using (var state = ArrangeParserAndParse(inputCode))
            {

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }
    }
}