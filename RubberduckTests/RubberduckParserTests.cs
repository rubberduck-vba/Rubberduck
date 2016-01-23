using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests
{
    [TestClass]
    public class RubberduckParserTests
    {
        [TestMethod]
        public void ParseResultDeclarations_IncludeVbaStandardLibDeclarations()
        {
            Assert.Fail();
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                                 .Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var state = new RubberduckParserState();
            var vbe = builder.AddProject(project).Build();
            var parser = new RubberduckParser(vbe.Object, state);

            //Act
            parser.ParseComponent(project.Object.VBComponents.Cast<VBComponent>().First());

            //Assert
            Assert.IsTrue(state.AllDeclarations.Any(item => item.IsBuiltIn));
        }

        //[TestMethod]
        //public void ParseResultDeclarations_MockHost_ExcludeExcelDeclarations()
        //{
        //    //Arrange
        //    var builder = new MockVbeBuilder();
        //    var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
        //        .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
        //        .Build().Object;

        //    var codePaneFactory = new CodePaneWrapperFactory();
        //    var mockHost = new Mock<IHostApplication>();
        //    mockHost.SetupAllProperties();

        //    //Act
        //    var parseResult = new RubberduckParser().Parse(project);

        //    //Assert
        //    Assert.IsFalse(parseResult.Declarations.Items.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        //}

        //[TestMethod]
        //public void ParseResultDeclarations_ExcelHost_IncludesExcelDeclarations()
        //{
        //    //Arrange
        //    var builder = new MockVbeBuilder();
        //    var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
        //        .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
        //        .AddReference("Excel", @"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE", true)
        //        .Build();
        //    var vbe = builder.AddProject(project).Build();

        //    var codePaneFactory = new CodePaneWrapperFactory();

        //    //Act
        //    var parseResult = new RubberduckParser().Parse(project.Object);

        //    //Assert
        //    Assert.IsTrue(parseResult.Declarations.Items.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        //}
    }
}
