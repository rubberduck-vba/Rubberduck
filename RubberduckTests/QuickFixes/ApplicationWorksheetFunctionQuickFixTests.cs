using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ApplicationWorksheetFunctionQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
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

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new ApplicationWorksheetFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(project.Object.VBComponents.First()).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
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

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new ApplicationWorksheetFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(project.Object.VBComponents.First()).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new ApplicationWorksheetFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(project.Object.VBComponents.First()).GetText());
            }
        }
    }
}
