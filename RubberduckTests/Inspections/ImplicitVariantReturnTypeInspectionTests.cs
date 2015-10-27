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
    public class ImplicitVariantReturnTypeInspectionTests
    {
        [TestMethod]
        public void ImplicitVariantReturnType_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo()
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitVariantReturnType_ReturnsResult_PropertyGet()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitVariantReturnType_ReturnsResult_MultipleFunctions()
        {
            const string inputCode =
@"Function Foo()
End Function

Function Goo()
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitVariantReturnType_DoesNotReturnResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitVariantReturnType_ReturnsResult_MultipleSubs_SomeReturning()
        {
            const string inputCode =
@"Function Foo()
End Function

Function Goo() As String
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser().Parse(project);

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitVariantReturnType_QuickFixWorks_Function()
        {
            const string inputCode =
@"Function Foo()
End Function";

            const string expectedCode =
@"Function Foo() As Variant
End Function";

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

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ImplicitVariantReturnType_QuickFixWorks_PropertyGet()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            const string expectedCode =
@"Property Get Foo() As Variant
End Property";

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

            var inspection = new ImplicitVariantReturnTypeInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ImplicitVariantReturnTypeInspection();
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitVariantReturnTypeInspection";
            var inspection = new ImplicitVariantReturnTypeInspection();

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}