using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.ToDoItems;
using RubberduckTests.Mocks;
using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers;

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
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, content);

            var vbe = builder.AddProject(project.Build()).Build();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            var vm = new ToDoExplorerViewModel(parser.State, GetConfigService(), GetOperatingSystemMock().Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var comments = vm.Items.Select(s => s.Type);

            Assert.IsTrue(comments.SequenceEqual(new[] { "TODO ", "NOTE ", "BUG " }));
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
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, content);

            var vbe = builder.AddProject(project.Build()).Build();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            var vm = new ToDoExplorerViewModel(parser.State, GetConfigService(), GetOperatingSystemMock().Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var comments = vm.Items.Select(s => s.Type);

            Assert.IsTrue(comments.SequenceEqual(new[] { "TODO ", "NOTE ", "BUG ", "BUG " }));
        }

        [TestMethod]
        public void RemoveRemovesComment()
        {
            var input =
@"Dim d As Variant  ' bug should be Integer";

            var expected =
@"Dim d As Variant  ";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, input)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            var vm = new ToDoExplorerViewModel(parser.State, GetConfigService(), GetOperatingSystemMock().Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            vm.SelectedItem = vm.Items.Single();
            vm.RemoveCommand.Execute(null);

            var module = project.Object.VBComponents[0].CodeModule;
            Assert.AreEqual(expected, module.Content());
            Assert.IsFalse(vm.Items.Any());
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

            var userSettings = new UserSettings(null, null, todoSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        private Mock<IOperatingSystem> GetOperatingSystemMock()
        {
            return new Mock<IOperatingSystem>();
        }
    }
}
