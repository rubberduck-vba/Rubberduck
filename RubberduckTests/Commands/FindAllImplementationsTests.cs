using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class FindAllImplementationsTests
    {
        [TestCategory("Commands")]
        [TestMethod]
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
        }

        [TestCategory("Commands")]
        [TestMethod]
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.First(s => s.IdentifierName == "IClass1_Foo"));

            Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
        }

        [TestCategory("Commands")]
        [TestMethod]
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

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(null);

            Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllImplementations_NoResults_DisplayMessageBox()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out component, default(Selection));
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, messageBox.Object, state, vbe.Object, vm, null);

            command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [TestCategory("Commands")]
        [TestMethod]
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var navigateCommand = new Mock<INavigateCommand>();

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(navigateCommand.Object, null, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllImplementations_NullTarget_Aborts()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, state, vbe.Object, vm, null);

            command.Execute(null);

            Assert.IsFalse(vm.Tabs.Any());
        }

        [TestCategory("Commands")]
        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, state, vbe.Object, vm, null);

            command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            Assert.IsFalse(vm.Tabs.Any());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllImplementations_CanExecute_NullTarget()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, state, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, state, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo")));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllImplementations_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllImplementationsCommand(null, null, state, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(null));
        }
    }
}
