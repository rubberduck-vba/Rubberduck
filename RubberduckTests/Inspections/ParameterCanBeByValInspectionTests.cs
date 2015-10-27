using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ParameterCanBeByValInspectionTests
    {
        [TestMethod]
        public void ParameterCanByByVal_ReturnsResult_PassedByNotSpecified()
        {
            const string inputCode =
@"Sub Foo(arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ParameterCanByByVal_ReturnsResult_PassedByRef_Unassigned()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ParameterCanByByVal_ReturnsResult_Multiple()
        {
            const string inputCode =
@"Sub Foo(arg1 As String, arg2 As Date)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ParameterCanByByVal_DoesNotReturnResult_PassedByValExplicitly()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ParameterCanByByVal_DoesNotReturnResult_PassedByRefAndAssigned()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ParameterCanByByVal_ReturnsResult_SomeParams()
        {
            const string inputCode =
@"Sub Foo(arg1 As String, ByVal arg2 As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ParameterCanByByVal_ReturnsResult_QuickFixWorks_PassedByUnspecified()
        {
            const string inputCode =
@"Sub Foo(arg1 As String)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            inspection.GetInspectionResults(parseResult).First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ParameterCanByByVal_ReturnsResult_QuickFixWorks_PassedByRefUnassigned()
        {
            const string inputCode =
@"Sub Foo(arg1 As String)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal arg1 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ParameterCanBeByValInspection();
            inspection.GetInspectionResults(parseResult).First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ParameterCanBeByValInspection();
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ParameterCanBeByValInspection";
            var inspection = new ParameterCanBeByValInspection();

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}