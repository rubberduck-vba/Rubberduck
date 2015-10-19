using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests
{
    [TestClass]
    public class RubberduckParserTests
    {
        [TestMethod]
        public void ParseResultDeclarations_IncludeVbaStandardLibDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            
            //Act
            var parseResult = new RubberduckParser(codePaneFactory, project.VBE).Parse(project);

            //Assert
            Assert.IsTrue(parseResult.Declarations.Items.Any(item => item.IsBuiltIn));
        }

        [TestMethod]
        public void ParseResultDeclarations_MockHost_ExcludeExcelDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .Build().Object;

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            //Act
            var parseResult = new RubberduckParser(codePaneFactory, project.VBE).Parse(project);

            //Assert
            Assert.IsFalse(parseResult.Declarations.Items.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        }

        [TestMethod]
        public void ParseResultDeclarations_ExcelHost_IncludesExcelDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .AddReference("Excel", @"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var codePaneFactory = new CodePaneWrapperFactory();

            //Act
            var parseResult = new RubberduckParser(codePaneFactory, vbe.Object).Parse(project.Object);

            //Assert
            Assert.IsTrue(parseResult.Declarations.Items.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        }
    }
}
