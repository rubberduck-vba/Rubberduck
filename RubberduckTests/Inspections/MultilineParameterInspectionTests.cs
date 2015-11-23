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
//    public class MultilineParameterInspectionTests
//    {
//        [TestMethod]
//        public void MultilineParameter_ReturnsResult()
//        {
//            const string inputCode = 
//@"Public Sub Foo(ByVal _
//    Var1 _
//    As _
//    Integer)
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

//            var inspection = new MultilineParameterInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultilineParameter_DoesNotReturnResult()
//        {
//            const string inputCode =
//@"Public Sub Foo(ByVal Var1 As Integer)
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

//            var inspection = new MultilineParameterInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(0, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultilineParameter_ReturnsMultipleResults()
//        {
//            const string inputCode =
//@"Public Sub Foo( _
//    ByVal _
//    Var1 _
//    As _
//    Integer, _
//    ByVal _
//    Var2 _
//    As _
//    Date)
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

//            var inspection = new MultilineParameterInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(2, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultilineParameter_ReturnsResults_SomeParams()
//        {
//            const string inputCode =
//@"Public Sub Foo(ByVal _
//    Var1 _
//    As _
//    Integer, ByVal Var2 As Date)
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

//            var inspection = new MultilineParameterInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void MultilineParameter_QuickFixWorks()
//        {
//            const string inputCode =
//@"Public Sub Foo( _
//    ByVal _
//    Var1 _
//    As _
//    Integer)
//End Sub";

//            const string expectedCode =
//@"Public Sub Foo( _
//    ByVal Var1 As Integer)
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

//            var inspection = new MultilineParameterInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            inspectionResults.First().QuickFixes.First().Fix();

//            Assert.AreEqual(expectedCode, module.Lines());
//        }

//        [TestMethod]
//        public void InspectionType()
//        {
//            var inspection = new MultilineParameterInspection();
//            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
//        }

//        [TestMethod]
//        public void InspectionName()
//        {
//            const string inspectionName = "MultilineParameterInspection";
//            var inspection = new MultilineParameterInspection();

//            Assert.AreEqual(inspectionName, inspection.Name);
//        }
//    }
}