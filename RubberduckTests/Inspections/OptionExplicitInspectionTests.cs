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
    public class OptionExplicitInspectionTests
    {
        [TestMethod]
        public void NotAlreadySpecified_ReturnsResult()
        {
            const string inputCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new OptionExplicitInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void AlreadySpecified_DoesNotReturnResult()
        {
            const string inputCode = @"Option Explicit";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new OptionExplicitInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void NotAlreadySpecified_ReturnsMultipleResults()
        {
            const string inputCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new OptionExplicitInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void PartiallySpecified_ReturnsResults()
        {
            const string inputCode1 = @"";
            const string inputCode2 = @"Option Explicit";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new OptionExplicitInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void NotAlreadySpecified_QuickFixWorks()
        {
            const string inputCode = @"";
            const string expectedCode =
@"Option Explicit

";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build().Object;
            var module = project.VBComponents.Item(0).CodeModule;

            var codePaneFactory = new CodePaneWrapperFactory();
            var parseResult = new RubberduckParser(codePaneFactory).Parse(project);

            var inspection = new OptionExplicitInspection();
            var inspectionResults = inspection.GetInspectionResults(parseResult);

            inspectionResults.First().QuickFixes.First().Fix();
            
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new OptionExplicitInspection();
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "OptionExplicitInspection";
            var inspection = new OptionExplicitInspection();

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}