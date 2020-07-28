using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ApplicationWorksheetFunctionQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ApplicationWorksheetFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ApplicationWorksheetFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ApplicationWorksheetFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ApplicationWorksheetFunctionQuickFix();
        }

        protected override IVBE TestVbe(string code, out IVBComponent component)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference(ReferenceLibrary.Excel)
                .Build();

            var vbe = builder.AddProject(project).Build().Object;
            component = project.Object.VBComponents[0];
            return vbe;
        }
    }
}
