using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ProcedureNotUsedInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Private Sub Foo()
    Goo
End Sub

Private Sub Goo()
    Foo
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_ReturnsResult_SomeProceduresUsed()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
    Foo
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_DoesNotReturnResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_HandlerIsIgnoredForUnraisedEvent()
        {
            //Input
            const string inputCode1 = @"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count(result => result.Target.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_NoResultForClassInitialize()
        {
            //Input
            const string inputCode =
@"Private Sub Class_Initialize()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_NoResultForCasedClassInitialize()
        {
            //Input
            const string inputCode =
@"Private Sub class_initialize()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_NoResultForClassTerminate()
        {
            //Input
            const string inputCode =
@"Private Sub Class_Terminate()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_NoResultForCasedClassTerminate()
        {
            //Input
            const string inputCode =
@"Private Sub class_terminate()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }


        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ProcedureNotUsed
Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode = @"";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new RemoveUnusedDeclarationQuickFix(state).Fix(inspectionResults.First());

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProcedureNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode =
@"'@Ignore ProcedureNotUsed
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ProcedureNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ProcedureNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ProcedureNotUsedInspection";
            var inspection = new ProcedureNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
