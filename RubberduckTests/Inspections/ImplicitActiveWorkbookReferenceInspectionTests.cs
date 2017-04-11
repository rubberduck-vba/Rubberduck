using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ImplicitActiveWorkbookReferenceInspectionTests
    {
        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
        public void ImplicitActiveWorkbookReference_ReportsWorksheets()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();


            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveWorkbookReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        [TestCategory("Inspections")]
        public void ImplicitActiveWorkbookReference_Ignored_DoesNotReportRange()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet

    '@Ignore ImplicitActiveWorkbookReference
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var vbe = builder.AddProject(project).Build();


            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveWorkbookReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"TestFiles\")]
        public void ImplicitActiveWorkbookReference_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            const string expectedCode =
                @"
Sub foo()
    Dim sheet As Worksheet
'@Ignore ImplicitActiveWorkbookReference
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();


            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitActiveWorkbookReferenceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(parser.State, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, parser.State.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ImplicitActiveWorkbookReferenceInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitActiveWorkbookReferenceInspection";
            var inspection = new ImplicitActiveWorkbookReferenceInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
