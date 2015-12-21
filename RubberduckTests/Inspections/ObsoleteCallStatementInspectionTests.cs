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
//    public class ObsoleteCallStatementInspectionTests
//    {
//        [TestMethod]
//        public void ObsoleteCallStatement_ReturnsResult()
//        {
//            const string inputCode = 
//@"Sub Foo()
//    Call Foo
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

//            var inspection = new ObsoleteCallStatementInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteCallStatement_DoesNotReturnResult()
//        {
//            const string inputCode =
//@"Sub Foo()
//    Foo
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

//            var inspection = new ObsoleteCallStatementInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(0, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteCallStatement_ReturnsMultipleResults()
//        {
//            const string inputCode =
//@"Sub Foo()
//    Call Goo(1, ""test"")
//End Sub
//
//Sub Goo(arg1 As Integer, arg1 As String)
//    Call Foo
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

//            var inspection = new ObsoleteCallStatementInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(2, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteCallStatement_ReturnsResults_SomeObsoleteCallStatements()
//        {
//            const string inputCode =
//@"Sub Foo()
//    Call Goo(1, ""test"")
//End Sub
//
//Sub Goo(arg1 As Integer, arg1 As String)
//    Foo
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

//            var inspection = new ObsoleteCallStatementInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            Assert.AreEqual(1, inspectionResults.Count());
//        }

//        [TestMethod]
//        public void ObsoleteCallStatement_QuickFixWorks_RemoveCallStatement()
//        {
//            const string inputCode =
//@"Sub Foo()
//    Call Goo(1, ""test"")
//End Sub
//
//Sub Goo(arg1 As Integer, arg1 As String)
//    Call Foo
//End Sub";

//            const string expectedCode =
//@"Sub Foo()
//    Goo 1, ""test""
//End Sub
//
//Sub Goo(arg1 As Integer, arg1 As String)
//    Foo
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

//            var inspection = new ObsoleteCallStatementInspection();
//            var inspectionResults = inspection.GetInspectionResults(parseResult);

//            foreach (var inspectionResult in inspectionResults)
//            {
//                inspectionResult.QuickFixes.First().Fix();
//            }

//            var actual = module.Lines();
//            Assert.AreEqual(expectedCode, actual);
//        }

//        [TestMethod]
//        public void InspectionType()
//        {
//            var inspection = new ObsoleteCallStatementInspection();
//            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
//        }

//        [TestMethod]
//        public void InspectionName()
//        {
//            const string inspectionName = "ObsoleteCallStatementInspection";
//            var inspection = new ObsoleteCallStatementInspection();

//            Assert.AreEqual(inspectionName, inspection.Name);
//        }
//    }
}