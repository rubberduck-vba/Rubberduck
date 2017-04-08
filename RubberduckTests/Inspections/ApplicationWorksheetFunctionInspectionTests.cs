using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ApplicationWorksheetFunctionInspectionTests
    {
        private static ParseCoordinator ArrangeParser(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.State.AddTestLibrary("Excel.1.8.xml");
            return parser;
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_ReturnsResult_GlobalApplication()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub
";

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
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

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
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

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
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

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_DoesNotReturnResult_ExplicitUseGlobalApplication()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    foo = Application.WorksheetFunction.Pi
End Sub
";

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
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

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    '@Ignore ApplicationWorksheetFunction
    foo = Application.Pi
End Sub
";

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub
";

            const string expectedCode =
@"Sub ExcelSub()
    Dim foo As Double
'@Ignore ApplicationWorksheetFunction
    foo = Application.Pi
End Sub
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var component = vbe.Object.SelectedVBComponent;

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(parser.State, new[] {inspection}).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, parser.State.GetRewriter(component).GetText());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_UseExplicitlyQuickFixWorks()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub
";

            const string expectedCode =
@"Sub ExcelSub()
    Dim foo As Double
    foo = Application.WorksheetFunction.Pi
End Sub
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();
            
            new ApplicationWorksheetFunctionQuickFix(parser.State).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, parser.State.GetRewriter(project.Object.VBComponents.First()).GetText());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_UseExplicitlyQuickFixWorks_WithBlock()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    With Application
        foo = .Pi
    End With
End Sub
";

            const string expectedCode =
@"Sub ExcelSub()
    Dim foo As Double
    With Application
        foo = .WorksheetFunction.Pi
    End With
End Sub
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            new ApplicationWorksheetFunctionQuickFix(parser.State).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, parser.State.GetRewriter(project.Object.VBComponents.First()).GetText());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ApplicationWorksheetFunction_UseExplicitlyQuickFixWorks_HasParameters()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As String
    foo = Application.Proper(""foobar"")
End Sub
";

            const string expectedCode =
@"Sub ExcelSub()
    Dim foo As String
    foo = Application.WorksheetFunction.Proper(""foobar"")
End Sub
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            new ApplicationWorksheetFunctionQuickFix(parser.State).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, parser.State.GetRewriter(project.Object.VBComponents.First()).GetText());
        }
    }
}