using System.Linq;
using System.Windows.Forms;
using NUnit.Framework;
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
    [TestFixture]
    public class FindAllReferencesTests
    {
        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                Assert.AreEqual(4, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                command.Execute(null);

                Assert.AreEqual(4, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_NoResults_DisplayMessageBox()
        {
            const string inputCode =
                @"Public Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

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
        }

        [Category("Commands")]
        [Test]
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
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var navigateCommand = new Mock<INavigateCommand>();

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(navigateCommand.Object, null, state, vbe.Object, vm, null);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_NullTarget_Aborts()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                command.Execute(null);

                Assert.IsFalse(vm.Tabs.Any());
            }
        }

        [Category("Commands")]
        [Test]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo"));

                Assert.IsFalse(vm.Tabs.Any());
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_CanExecute_NullTarget()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                Assert.IsFalse(command.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                state.SetStatusAndFireStateChanged(this, ParserState.ResolvedDeclarations);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                Assert.IsFalse(command.CanExecute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Foo")));
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out component);
            vbe.Setup(s => s.ActiveCodePane).Returns(value: null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);

                Assert.IsFalse(command.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_ControlMultipleResults_ReturnsCorrectNumber()
        {
            var code = @"
Public Sub DoSomething()
    TextBox1.Height = 20
    TextBox1.Width = 200
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);
                var target = state.AllUserDeclarations.Single(s => s.IdentifierName == "TextBox1");

                command.Execute(target);

                Assert.AreEqual(1, vm.Tabs.Count);
                Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_ControlSingleResult_Navigates()
        {
            var code = @"
Public Sub DoSomething()
    TextBox1.Height = 20
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);
                var navigateCommand = new Mock<INavigateCommand>();

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(navigateCommand.Object, null, state, vbe.Object, vm, null);
                var target = state.AllUserDeclarations.Single(s => s.IdentifierName == "TextBox1");

                command.Execute(target);

                navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_ControlNoResults_DisplaysMessageBox()
        {
            var code = @"
Public Sub DoSomething()
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, messageBox.Object, state, vbe.Object, vm, null);
                var target = state.AllUserDeclarations.Single(s => s.IdentifierName == "TextBox1");

                command.Execute(target);

                messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                    It.IsAny<MessageBoxIcon>()), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_ControlMultipleSelection_IsNotEnabled()
        {
            var code = @"
Public Sub DoSomething()
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").AddControl("TextBox2").Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);

                var targets = state.AllUserDeclarations.Where(s => s.IdentifierName.StartsWith("TextBox"));
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, null, null);


                Assert.IsFalse(command.CanExecute(targets));
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_FormMultipleResults_ReturnsCorrectNumber()
        {
            var code = @"
Public Sub DoSomething()
    Form1.Width = 20
    Form1.Height = 200
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, null, state, vbe.Object, vm, null);
                var target = state.AllUserDeclarations.Single(s => s.IdentifierName == "Form1");

                command.Execute(target);

                Assert.AreEqual(1, vm.Tabs.Count);
                Assert.AreEqual(2, vm.Tabs[0].SearchResults.Count);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_FormSingleResult_Navigates()
        {
            var code = @"
Public Sub DoSomething()
    Form1.Height = 20
End Sub
";
            var navigateCommand = new Mock<INavigateCommand>();
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);


            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(navigateCommand.Object, null, state, vbe.Object, vm, null);

                command.Execute(state.AllUserDeclarations.Single(s => s.IdentifierName == "Form1"));

                navigateCommand.Verify(n => n.Execute(It.IsAny<object>()), Times.Once);
            }
        }

        [Category("Commands")]
        [Test]
        public void FindAllReferences_FormNoResults_DisplaysMessageBox()
        {
            var code = @"
Public Sub DoSomething()
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).Build();

            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();
            vbe.SetupGet(v => v.SelectedVBComponent).Returns(form.Object);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                AssertParserReady(state);

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

                var vm = new SearchResultsWindowViewModel();
                var command = new FindAllReferencesCommand(null, messageBox.Object, state, vbe.Object, vm, null);
                var target = state.AllUserDeclarations.Single(s => s.IdentifierName == "Form1");

                command.Execute(target);

                messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                    It.IsAny<MessageBoxIcon>()), Times.Once);
            }
        }

        private void AssertParserReady(RubberduckParserState state)
        {
            if (state.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }
            if (state.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }
        }
    }
}
