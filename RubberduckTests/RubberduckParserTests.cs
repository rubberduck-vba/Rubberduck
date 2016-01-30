using System.Linq;
using System.Threading;
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
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(0, 1);

        void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Ready)
            {
                _semaphore.Release();
            }
        }

        [TestMethod, Timeout(1000)]
        public void ParseResultDeclarations_IncludeVbaStandardLibDeclarations()
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

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            //Act
            parser.ParseComponent(project.Object.VBComponents.Cast<VBComponent>().First());

            //Assert
            Assert.IsTrue(parser.State.AllDeclarations.Any(item => item.IsBuiltIn));
        }

        [TestMethod, Timeout(1000)]
        public void ParseResultDeclarations_MockHost_ExcludeExcelDeclarations()
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
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            //Assert
            Assert.IsFalse(parseResult.State.AllDeclarations.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        }

        [TestMethod, Timeout(1000)]
        public void ParseResultDeclarations_ExcelHost_IncludesExcelDeclarations()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, "")
                .AddReference("Excel", @"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE", true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            //Act
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            //Assert
            Assert.IsTrue(parseResult.State.AllDeclarations.Any(item => item.IsBuiltIn && item.ParentScope.StartsWith("Excel")));
        }
    }
}
