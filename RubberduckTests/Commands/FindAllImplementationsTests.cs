using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Interaction.Navigation;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class FindAllImplementationsTests
    {
        [Category("Commands")]
        [Test]
        public void FindAllImplementations_ReturnsCorrectNumber()
        {
            const string inputClass =
                @"Implements IClass1

Public Sub IClass1_Foo()
End Sub";

            const string inputInterface =
                @"Public Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputClass)
                .AddComponent("Class2", ComponentType.ClassModule, inputClass)
                .AddComponent("IClass1", ComponentType.ClassModule, inputInterface)
                .Build();

            var vbe = builder.AddProject(project).Build();
            
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_SelectedImplementation_ReturnsCorrectNumber()
        {
            const string inputClass =
                @"Implements IClass1

Public Sub IClass1_Foo()
End Sub";

            const string inputInterface =
                @"Public Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputClass)
                .AddComponent("Class2", ComponentType.ClassModule, inputClass)
                .AddComponent("IClass1", ComponentType.ClassModule, inputInterface)
                .Build();

            var vbe = builder.AddProject(project).Build();
            
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(state.AllUserDeclarations.First(s => s.IdentifierName == "IClass1_Foo"));

                Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_ReturnsCorrectNumberForProperty()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Private Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Private Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, implementation)
                .AddComponent("Class2", ComponentType.ClassModule, implementation)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface)
                .Build();

            var vbe = builder.AddProject(project).Build();
            
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo" && s.DeclarationType == DeclarationType.PropertyGet));

                Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_SelectedReference_ReturnsCorrectNumber()
        {
            const string inputClass =
                @"Implements IClass1

Public Sub IClass1_Foo()
End Sub

Public Sub Buzz()
    IClass1_Foo
End Sub";

            const string inputInterface =
                @"Public Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputClass, new Selection(7, 5, 7, 5))
                .AddComponent("Class2", ComponentType.ClassModule, inputClass)
                .AddComponent("IClass1", ComponentType.ClassModule, inputInterface)
                .Build();

            var vbe = builder.AddProject(project).Build();
            vbe.Setup(v => v.ActiveCodePane).Returns(project.Object.VBComponents["Class1"].CodeModule.CodePane);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(null);

                Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_NoResults_DisplayMessageBox()
        {
            const string inputCode =
                @"Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var messageBox = new Mock<IMessageBox>();
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm, messageBox: messageBox.Object);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_SingleResult_Navigates()
        {
            const string inputClass =
                @"Implements IClass1

Public Sub IClass1_Foo()
End Sub";

            const string inputInterface =
                @"Public Sub Foo()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputClass)
                .AddComponent("IClass1", ComponentType.ClassModule, inputInterface)
                .Build();

            var vbe = builder.AddProject(project).Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var navigateCommand = new Mock<INavigateCommand>();

                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm, navigateCommand.Object);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_NullTarget_Aborts()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(null);

                Assert.IsFalse(vm.Tabs.Any());
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_StateNotReady_Aborts()
        {
            const string inputCode =
                @"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo: Foo
    Foo
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations, CancellationToken.None);

                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                Assert.IsFalse(vm.Tabs.Any());
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_CanExecute_NullTarget()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                Assert.IsFalse(command.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_CanExecute_StateNotReady()
        {
            const string inputCode =
                @"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo: Foo
    Foo
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations, CancellationToken.None);

                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                Assert.IsFalse(command.CanExecute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo")));
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllImplementations_CanExecute_NullActiveCodePane()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var vm = ArrangeSearchResultsWindowViewModel();
                var service = ArrangeFindAllImplementationsService(state, vm);
                var command = ArrangeFindAllImplementationsCommand(state, vbe, vm, service);

                Assert.IsFalse(command.CanExecute(null));
            }
        }

        private static SearchResultsWindowViewModel ArrangeSearchResultsWindowViewModel()
        {
            return new SearchResultsWindowViewModel();
        }

        private static FindAllImplementationsService ArrangeFindAllImplementationsService(RubberduckParserState state,
            ISearchResultsWindowViewModel viewModel, INavigateCommand navigateCommand = null, IMessageBox messageBox = null,
            SearchResultPresenterInstanceManager presenterService = null, IUiDispatcher uiDispatcher = null)
        {
            return new FindAllImplementationsService(
                navigateCommand ?? new Mock<INavigateCommand>().Object,
                messageBox ?? new Mock<IMessageBox>().Object,
                state,
                viewModel,
                presenterService,
                uiDispatcher ?? new Mock<IUiDispatcher>().Object);
        }

        private static FindAllImplementationsCommand ArrangeFindAllImplementationsCommand(RubberduckParserState state,
            Mock<IVBE> vbe)
        {
            var viewModel = ArrangeSearchResultsWindowViewModel();
            var finder = ArrangeFindAllImplementationsService(state, viewModel);
            return ArrangeFindAllImplementationsCommand(state, vbe, viewModel, finder);
        }

        private static FindAllImplementationsCommand ArrangeFindAllImplementationsCommand(RubberduckParserState state,
            Mock<IVBE> vbe, ISearchResultsWindowViewModel viewModel, FindAllImplementationsService finder)
        {
            return ArrangeFindAllImplementationsCommand(state, vbe, viewModel, finder, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        private static FindAllImplementationsCommand ArrangeFindAllImplementationsCommand(RubberduckParserState state,
            Mock<IVBE> vbe, ISearchResultsWindowViewModel viewModel, FindAllImplementationsService finder,
            Mock<IVbeEvents> vbeEvents)
        {
            var selectionService = new SelectionService(vbe.Object, state.ProjectsProvider);
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            return new FindAllImplementationsCommand(state, selectedDeclarationService, viewModel, finder, vbeEvents.Object);
        }
    }
}
