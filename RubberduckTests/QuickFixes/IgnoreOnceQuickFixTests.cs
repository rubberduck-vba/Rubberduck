using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Inspections;

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

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AssignedByValParameter_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);

            var state = MockParser.CreateAndParse(vbe.Object);
            var inspection = new AssignedByValParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ConstantNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
@"Public Sub Foo()
'@Ignore ConstantNotUsed
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void EmptyStringLiteral_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo(ByRef arg1 As String)
    arg1 = """"
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByRef arg1 As String)
'@Ignore EmptyStringLiteral
    arg1 = """"
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyStringLiteralInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void EncapsulatePublicField_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public fizz As Boolean";

            const string expectedCode =
@"'@Ignore EncapsulatePublicField
Public fizz As Boolean";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EncapsulatePublicFieldInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        [TestCategory("Unused Value")]
        public void FunctionReturnValueNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            const string expectedCode =
@"'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new FunctionReturnValueNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void AnnotationListFollowedByCommentAddsAnnotationCorrectly()
        {
            const string inputCode = @"
Public Function GetSomething() As Long
    '@Ignore VariableNotAssigned: Is followed by a comment.
    Dim foo
    GetSomething = foo
End Function
";

            const string expectedCode = @"
Public Function GetSomething() As Long
    '@Ignore VariableTypeNotDeclared, VariableNotAssigned: Is followed by a comment.
    Dim foo
    GetSomething = foo
End Function
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }


    }
}
