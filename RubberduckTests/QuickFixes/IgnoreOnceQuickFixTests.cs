using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class IgnoreOnceQuickFixTests
    {

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("QuickFixes")]
        public void ApplicationWorksheetFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub";

            const string expectedCode =
@"Sub ExcelSub()
    Dim foo As Double
'@Ignore ApplicationWorksheetFunction
    foo = Application.Pi
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var component = vbe.Object.SelectedVBComponent;

            var parser = MockParser.Create(vbe.Object);

            parser.State.AddTestLibrary("Excel.1.8.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ApplicationWorksheetFunctionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(parser.State, new[] { inspection }).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, parser.State.GetRewriter(component).GetText());
        }


    }
}
