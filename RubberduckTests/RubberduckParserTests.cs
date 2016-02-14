using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests
{
    [TestClass]
    public class RubberduckParserTests
    {
        /// <summary>
        /// Built-in declarations are included in the parser state explicitly at startup.
        /// </summary>
        [TestMethod]
        public void parserDeclarations_ExcludeBuiltInDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                                 .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                                 .Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var vbe = builder.AddProject(project).Build();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            //Act
            parser.ParseComponent(project.Object.VBComponents.Cast<VBComponent>().First());

            //Assert
            Assert.IsFalse(parser.State.AllDeclarations.Any(item => item.IsBuiltIn));
        }

        [TestMethod]
        public void parserDeclarations_MockHost_ExcludeExcelDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .Build();
            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            //Act
            var state = new RubberduckParserState();
            state.AddBuiltInDeclarations(mockHost.Object);
            var parser = new RubberduckParser(vbe.Object, state);

            //Assert
            Assert.IsFalse(parser.State.AllDeclarations.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        }

        [TestMethod]
        public void parserDeclarations_ExcelHost_IncludesExcelDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .AddReference("Excel", @"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var host = vbe.Object.HostApplication();

            //Act
            var state =  new RubberduckParserState();
            state.AddBuiltInDeclarations(host);
            var parser = new RubberduckParser(vbe.Object, state);

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            //Assert
            Assert.IsTrue(parser.State.AllDeclarations.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        }
    }
}
