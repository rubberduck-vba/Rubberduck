using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ProcedureNotUsedInspectionTests
    {
        [TestMethod]
        public void ProcedureNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ProcedureNotUsed_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ProcedureUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Private Sub Foo()
    Goo
End Sub

Private Sub Goo()
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ProcedureNotUsed_ReturnsResult_SomeProceduresUsed()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
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

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ProcedureNotUsed_DoesNotReturnResult_EventImplementation()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ProcedureNotUsed_ReturnsResult_PublicModuleMember()
        {
            //Input
            const string inputCode =
@"Private Sub Class_Initialize()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ProcedureNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new RubberduckCodePaneFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ProcedureNotUsedInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            inspectionResults.First().GetQuickFixes().First().Value();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ProcedureNotUsedInspection();
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ProcedureNotUsedInspection";
            var inspection = new ProcedureNotUsedInspection();

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}