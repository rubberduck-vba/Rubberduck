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
    public class ImplicitPublicMemberInspectionTests
    {
        [TestMethod]
        public void ImplicitPublicMember_ReturnsResult_Sub()
        {
            const string inputCode =
@"Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ImplicitPublicMemberInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitPublicMember_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo() As Boolean
    Foo = True
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ImplicitPublicMemberInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitPublicMember_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
@"Sub Foo()
End Sub

Sub Goo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ImplicitPublicMemberInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitPublicMember_DoesNotReturnResult()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ImplicitPublicMemberInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitPublicMember_ReturnsResult_SomeImplicitlyPublicSubs()
        {
            const string inputCode =
@"Private Sub Foo()
End Sub

Sub Goo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ImplicitPublicMemberInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ImplicitPublicMember_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 as Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new ImplicitPublicMemberInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ImplicitPublicMemberInspection();
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitPublicMemberInspection";
            var inspection = new ImplicitPublicMemberInspection();

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}