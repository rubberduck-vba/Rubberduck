using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.ToDoItems;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.TodoExplorer
{
    [TestClass]
    public class TodoExplorerTests
    {
        [TestMethod]
        public void PicksUpComments()
        {
            var content =
@"' Todo this is a todo comment
' Note this is a note comment
' Bug this is a bug comment
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, content);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var vm = new ToDoExplorerViewModel(parser.State, GetConfigService());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var comments = vm.Items.Select(s => s.Type);

            Assert.IsTrue(comments.SequenceEqual(new[] {"TODO ", "NOTE ", "BUG "}));
        }

        [TestMethod]
        public void PicksUpComments_StrangeCasing()
        {
            var content =
@"' tODO this is a todo comment
' NOTE  this is a note comment
' bug this is a bug comment
' bUg this is a bug comment
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, content);

            var vbe = builder.AddProject(project.Build()).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            var vm = new ToDoExplorerViewModel(parser.State, GetConfigService());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var comments = vm.Items.Select(s => s.Type);

            Assert.IsTrue(comments.SequenceEqual(new[] { "TODO ", "NOTE ", "BUG ", "BUG " }));
        }

        private IGeneralConfigService GetConfigService()
        {
            var configService = new Mock<IGeneralConfigService>();
            configService.Setup(c => c.LoadConfiguration()).Returns(GetTodoConfig);
            
            return configService.Object;
        }

        private Configuration GetTodoConfig()
        {
            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE "),
                    new ToDoMarker("TODO "),
                    new ToDoMarker("BUG ")
                }
            };

            var userSettings = new UserSettings(null, todoSettings, null, null, null);
            return new Configuration(userSettings);
        }
    }
}