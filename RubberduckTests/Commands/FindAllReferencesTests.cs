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
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class FindAllReferencesTests
    {
        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_ReturnsCorrectNumber()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo: Foo
    Foo
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            Assert.AreEqual(4, vm.Tabs[0].SearchResults.Count);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_ReferenceSelected_ReturnsCorrectNumber()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo: Foo
    Foo
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, new Selection(5, 5, 5, 5));
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(null);

            Assert.AreEqual(4, vm.Tabs[0].SearchResults.Count);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_NoResults_DisplayMessageBox()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, messageBox.Object, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_SingleResult_Navigates()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var navigateCommand = new Mock<INavigateCommand>();

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(navigateCommand.Object, null, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_NullTarget_Aborts()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(null);

            Assert.IsFalse(vm.Tabs.Any());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_StateNotReady_Aborts()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo: Foo
    Foo
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            parser.State.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            command.Execute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            Assert.IsFalse(vm.Tabs.Any());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_CanExecute_NullTarget()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_CanExecute_StateNotReady()
        {
            const string inputCode =
@"Public Sub Foo()
End Sub

Private Sub Bar()
    Foo: Foo
    Foo
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            parser.State.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(parser.State.AllUserDeclarations.Single(s => s.IdentifierName == "Foo")));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_CanExecute_NullActiveCodePane()
        {
            //Arrange
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, parser.State, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(null));
        }
    }
}
