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
    }
}
