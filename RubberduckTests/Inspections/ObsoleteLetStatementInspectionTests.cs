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
//    public class ObsoleteLetStatementInspectionTests
//    {
//        [TestMethod]
//        public void ObsoleteLetStatement_ReturnsResult()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer
//    Dim var2 As Integer
//    
//    Let var2 = var1
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

//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteLetStatement_ReturnsResult_MultipleLets()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer
//    Dim var2 As Integer
//    
//    Let var2 = var1
//    Let var1 = var2
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

//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(2, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteLetStatement_DoesNotReturnResult()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer
//    Dim var2 As Integer
//    
//    var2 = var1
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

//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(0, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteLetStatement_ReturnsResult_SomeConstantsUsed()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer
//    Dim var2 As Integer
//    
//    Let var2 = var1
//    var1 = var2
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

//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteLetStatement_QuickFixWorks()
//        {
//            const string inputCode =
//@"Public Sub Foo()
//    Dim var1 As Integer
//    Dim var2 As Integer
//    
//    Let var2 = var1
//End Sub";

//            const string expectedCode =
//@"Public Sub Foo()
//    Dim var1 As Integer
//    Dim var2 As Integer
//    
//    var2 = var1
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

//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
//            var inspectionResults = inspection.GetInspectionResults();

//            inspectionResults.First().QuickFixes.First().Fix();

//            Assert.AreEqual(expectedCode, module.Lines());
//        }

//        [TestMethod]
//        public void InspectionType()
//        {
//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
//            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
//        }

//        [TestMethod]
//        public void InspectionName()
//        {
//            const string inspectionName = "ObsoleteLetStatementInspection";
//            var inspection = new ObsoleteLetStatementInspection(parseResult.State);

//            Assert.AreEqual(inspectionName, inspection.Name);
//        }
//    }
}