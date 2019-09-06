using System;
using NUnit.Framework;
using Moq;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.UI.ToDoItems;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.UIContext;
using Rubberduck.SettingsProvider;
using Rubberduck.ToDoItems;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.TodoExplorer
{
    [TestFixture]
    public class TodoExplorerTests
    {
        [Test]
        [Category("Annotations")]
        public void PicksUpComments()
        {
            const string inputCode =
                @"' Todo this is a todo comment
' Note this is a note comment
' Bug this is a bug comment
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var cs = GetConfigService(new[] { "TODO", "NOTE", "BUG" });
                var vm = ArrangeViewModel(state, cs);

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var comments = vm.Items.OfType<ToDoItem>().Select(s => s.Type);

                Assert.IsTrue(comments.SequenceEqual(new[] { "TODO", "NOTE", "BUG" }));
            }
        }

        [Test]
        [Category("Annotations")]
        public void PicksUpComments_StrangeCasing()
        {
            const string inputCode =
                @"' tODO this is a todo comment
' NOTE  this is a note comment
' bug this is a bug comment
' bUg this is a bug comment
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var cs = GetConfigService(new[] { "TODO", "NOTE", "BUG" });
                var vm = ArrangeViewModel(state, cs);

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var comments = vm.Items.OfType<ToDoItem>().Select(s => s.Type);

                Assert.IsTrue(comments.SequenceEqual(new[] { "TODO", "NOTE", "BUG", "BUG" }));
            }
        }

        [Test]
        [Category("Annotations")]
        public void PicksUpComments_SpecialCharacters()
        {
            const string inputCode =
                @"' To-do - this is a todo comment
' N@TE this is a note comment
' bug this should work with a trailing space
' bug: this should not be seen due to the colon
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var cs = GetConfigService(new[] { "TO-DO", "N@TE", "BUG " });
                var vm = ArrangeViewModel(state, cs);

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var comments = vm.Items.OfType<ToDoItem>().Select(s => s.Type);

                Assert.IsTrue(comments.SequenceEqual(new[] { "TO-DO", "N@TE", "BUG " }));
            }
        }

        [Test]
        [Category("Annotations")]
        public void AvoidsFalsePositiveComments()
        {
            const string inputCode =
                @"' Todon't should not get picked up
' Debug.print() would trigger false positive if word boundaries not used
' Denoted 
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                var cs = GetConfigService(new[] { "TODO", "NOTE", "BUG" });
                var vm = ArrangeViewModel(state, cs);

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var comments = vm.Items.OfType<ToDoItem>().Select(s => s.Type);

                Assert.IsTrue(comments.Count() == 0);
            }
        }

        [Test]
        [Category("Annotations")]
        public void RemoveRemovesComment()
        {
            const string inputCode =
                @"Dim d As Variant  ' bug should be Integer";

            const string expected =
                @"Dim d As Variant  ";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var vbeEvents = MockVbeEvents.CreateMockVbeEvents(vbe);

            var parser = MockParser.Create(vbe.Object, vbeEvents:vbeEvents);
            using (var state = parser.State)
            {
                var cs = GetConfigService(new[] { "TODO", "NOTE", "BUG" });
                var vm = ArrangeViewModel(state, cs);
                
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                vm.SelectedItem = vm.Items.OfType<ToDoItem>().Single();
                vm.RemoveCommand.Execute(null);

                var module = project.Object.VBComponents[0].CodeModule;
                Assert.AreEqual(expected, module.Content());
            }
        }

        private IConfigurationService<Configuration> GetConfigService(string[] markers)
        {
            var configService = new Mock<IConfigurationService<Configuration>>();
            configService.Setup(c => c.Read()).Returns(GetTodoConfig(markers));

            return configService.Object;
        }

        private Configuration GetTodoConfig(string[] markers)
        {
            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = markers.Select(m => new ToDoMarker(m)).ToArray()
            };

            var userSettings = new UserSettings(null, null, null, todoSettings, null, null, null, null);
            return new Configuration(userSettings);
        }

        private IUiDispatcher GetMockedUiDispatcher()
        {
            var dispatcher = new Mock<IUiDispatcher>();
            dispatcher.Setup(m => m.Invoke(It.IsAny<Action>())).Callback((Action argument) => argument.Invoke());
            return dispatcher.Object;
        }

        private ToDoExplorerViewModel ArrangeViewModel(RubberduckParserState state, IConfigurationService<Configuration> configService)
        {
            return new ToDoExplorerViewModel(state, configService, null, GetMockedUiDispatcher(), null);
        }
    }
}
