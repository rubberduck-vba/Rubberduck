using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ProcedureShouldBeFunctionInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult()
        {
            const string inputCode =
@"Private Sub Foo(ByRef foo As Boolean)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
@"Private Sub Foo(ByRef foo As Boolean)
End Sub

Private Sub Goo(ByRef foo As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_Function()
        {
            const string inputCode =
@"Private Function Foo(ByRef bar As Integer) As Integer
    Foo = bar
End Function";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_SingleByValParam()
        {
            const string inputCode =
@"Private Sub Foo(ByVal foo As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnsResult_MultipleByValParams()
        {
            const string inputCode =
@"Private Sub Foo(ByVal foo As Integer, ByVal goo As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnsResult_MultipleByRefParams()
        {
            const string inputCode =
@"Private Sub Foo(ByRef foo As Integer, ByRef goo As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByRef a As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";
            
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_EventImplementation()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByRef arg1 As Integer)";
            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub Foo(ByRef foo As Boolean)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Integer
    Foo = arg1
End Function";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks_NoAsTypeClauseInParam()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1)
End Sub";

            const string expectedCode =
@"Private Function Foo(ByVal arg1) As Variant
    Foo = arg1
End Function";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks_DoesNotInterfereWithBody()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1 As Integer)
    arg1 = 6
    Goo arg1
End Sub

Sub Goo(ByVal a As Integer)
End Sub";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Integer
    arg1 = 6
    Goo arg1
    Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks_DoesNotInterfereWithBody_BodyOnSingleLine()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1 As Integer): arg1 = 6: Goo arg1: End Sub

Sub Goo(ByVal a As Integer)
End Sub";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Integer: arg1 = 6: Goo arg1:     Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks_DoesNotInterfereWithBody_BodyOnSingleLine_BodyHasStringLiteralContainingColon()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1 As Integer): arg1 = 6: Goo ""test: test"": End Sub

Sub Goo(ByVal a As String)
End Sub";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Integer: arg1 = 6: Goo ""test: test"":     Foo = arg1
End Function

Sub Goo(ByVal a As String)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks_UpdatesCallsAbove()
        {
            const string inputCode =
@"Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    Foo fizz
End Sub

Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    fizz = Foo(fizz)
End Sub

Private Function Foo(ByVal arg1 As Integer) As Integer
    Foo = arg1
End Function";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_QuickFixWorks_UpdatesCallsBelow()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1 As Integer)
End Sub

Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    Foo fizz
End Sub";

            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer) As Integer
    Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    fizz = Foo(fizz)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureShouldBeFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub Foo(ByRef arg1 As Integer)
End Sub";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ProcedureCanBeWrittenAsFunctionInspection";
            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
