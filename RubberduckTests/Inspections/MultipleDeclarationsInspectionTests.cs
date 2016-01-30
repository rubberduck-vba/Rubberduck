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
//    public class MultipleDeclarationsInspectionTests
//    {
//        [TestMethod]
//        public void MultipleDeclarations_ReturnsResult_Variables()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer, var2 As String
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_ReturnsResult_Constants()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Const var1 As Integer = 9, var2 As String = ""test""
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_ReturnsResult_StaticVariables()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Static var1 As Integer, var2 As String
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_ReturnsResult_MultipleDeclarations()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer, var2 As String
//    Dim var3 As Boolean, var4 As Date
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(2, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_ReturnsResult_SomeDeclarationsSeparate()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer, var2 As String
//    Dim var3 As Boolean
//End Sub";

//            //Arrange
//            var builder = new MockVbeBuilder();
//            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
//                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
//                .Build().Object;

//            var codePaneFactory = new CodePaneWrapperFactory();
//            var mockHost = new Mock<IHostApplication>();
//            mockHost.SetupAllProperties();
//            var parseResult = new RubberduckParser().Parse(project);

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_QuickFixWorks_Variables()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer, var2 As String
//End Sub";

//            const string expectedCode =
//@"Public Sub Foo()
//Dim var1 As Integer
//Dim var2 As String
//
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

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            inspectionResults.First().QuickFixes.First().Fix();

//            Assert.AreEqual(expectedCode, module.Lines());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_QuickFixWorks_Constants()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Const var1 As Integer = 9, var2 As String = ""test""
//End Sub";

//            const string expectedCode =
//@"Public Sub Foo()
//Const var1 As Integer = 9
//Const var2 As String = ""test""
//
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

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            inspectionResults.First().QuickFixes.First().Fix();

//            Assert.AreEqual(expectedCode, module.Lines());
//        }

//        [TestMethod]
//        public void MultipleDeclarations_QuickFixWorks_StaticVariables()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Static var1 As Integer, var2 As String
//End Sub";

//            const string expectedCode =
//@"Public Sub Foo()
//Static var1 As Integer
//Static var2 As String
//
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

//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            inspectionResults.First().QuickFixes.First().Fix();

//            Assert.AreEqual(expectedCode, module.Lines());
//        }

//        [TestMethod]
//        public void InspectionType()
//        {
//            var inspection = new MultipleDeclarationsInspection(parseResult.State);
//            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
//        }

//        [TestMethod]
//        public void InspectionName()
//        {
//            const string inspectionName = "MultipleDeclarationsInspection";
//            var inspection = new MultipleDeclarationsInspection(parseResult.State);

//            Assert.AreEqual(inspectionName, inspection.Name);
//        }
//    }
}