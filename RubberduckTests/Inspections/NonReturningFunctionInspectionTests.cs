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
//    [TestClass]
//    public class NonReturningFunctionInspectionTests
//    {
//        [TestMethod]
//        public void NonReturningFunction_ReturnsResult()
//        {
//            const string inputCode =
//@"Function Foo() As Boolean
//End Function";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void NonReturningFunction_ReturnsResult_MultipleFunctions()
//        {
//            const string inputCode =
//@"Function Foo() As Boolean
//End Function
//
//Function Goo() As String
//End Function";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(2, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void NonReturningFunction_DoesNotReturnResult()
//        {
//            const string inputCode =
//@"Function Foo() As Boolean
//    Foo = True
//End Function";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(0, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void NonReturningFunction_ReturnsResult_MultipleSubs_SomeReturning()
//        {
//            const string inputCode =
//@"Function Foo() As Boolean
//    Foo = True
//End Function
//
//Function Goo() As String
//End Function";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void NonReturningFunction_ReturnsResult_InterfaceImplementation()
//        {
//            //Input
//            const string inputCode1 =
//@"Function Foo() As Boolean
//End Function";
//            const string inputCode2 =
//@"Implements IClass1
//
//Function IClass1_Foo() As Boolean
//End Function";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void NonReturningFunction_QuickFixWorks()
//        {
//            const string inputCode =
//@"Function Foo() As Boolean
//End Function";

//            const string expectedCode =
//@"Sub Foo()
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;
//            var module = project.VBComponents.Item(0).CodeModule;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            inspectionResults.First().QuickFixes.First().Fix();

//            Assert.AreEqual(expectedCode, module.Lines());
//        }

//        [TestMethod]
//        public void NonReturningFunction_ReturnsResult_InterfaceImplementation_NoQuickFix()
//        {
//            //Input
//            const string inputCode1 =
//@"Function Foo() As Boolean
//End Function";
//            const string inputCode2 =
//@"Implements IClass1
//
//Function IClass1_Foo() As Boolean
//End Function";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(0, inspectionResults.First().QuickFixes.Count());
//        }

//        [TestMethod]
//        public void InspectionType()
//        {
//            var inspection = new NonReturningFunctionInspection(parseResult.State);
//            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
//        }

//        [TestMethod]
//        public void InspectionName()
//        {
//            const string inspectionName = "NonReturningFunctionInspection";
//            var inspection = new NonReturningFunctionInspection(parseResult.State);

//            Assert.AreEqual(inspectionName, inspection.Name);
//        }
//    }
}