using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

            command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, new Selection(5, 5, 5, 5));
            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, messageBox.Object, state, vbe.Object, vm, null);

            command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var navigateCommand = new Mock<INavigateCommand>();

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(navigateCommand.Object, null, state, vbe.Object, vm, null);

            command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_NullTarget_Aborts()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);
            state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

            command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

            Assert.IsFalse(vm.Tabs.Any());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_CanExecute_NullTarget()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo")));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void FindAllReferences_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var vm = new SearchResultsWindowViewModel();
            var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

            Assert.IsFalse(command.CanExecute(null));
        }
    }
}
