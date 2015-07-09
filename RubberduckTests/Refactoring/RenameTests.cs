using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RenameTests : RefactoringTestBase
    {
        [TestMethod]
        public void RenameRefactoring_RenameSub()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameVariable()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim val1 As Integer
End Sub";
            var selection = new Selection(2, 12, 2, 12); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim val2 As Integer
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameSub_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
    Foo
End Sub
";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Hoo()
End Sub

Private Sub Goo()
    Hoo
End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameVariable_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim val1 As Integer
    val1 = val1 + 5
End Sub";
            var selection = new Selection(2, 12, 2, 12); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim val2 As Integer
    val2 = val2 + 5
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
    arg1 = ""test""
End Sub";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
    arg2 = ""test""
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameGetterAndSetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) 
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Goo(ByVal arg1 As Integer) 
End Property

Private Property Set Goo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameGetterAndLetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo() 
End Property

Private Property Let Foo(ByVal arg1 As String) 
End Property";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Goo() 
End Property

Private Property Let Goo(ByVal arg1 As String) 
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 21, 1, 21); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Goo() As Boolean
    Goo = True
End Function";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFunction_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Foo = True
End Function

Private Sub Goo()
    Dim var1 As Boolean
    var1 = Foo()
End Sub
";
            var selection = new Selection(1, 21, 1, 21); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Hoo() As Boolean
    Hoo = True
End Function

Private Sub Goo()
    Dim var1 As Boolean
    var1 = Hoo()
End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RefactorWithDeclaration()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RenameModel(IDE.Object, parseResult, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(model.Target);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IRenamePresenter>> SetupFactory(RenameModel model)
        {
            var presenter = new Mock<IRenamePresenter>();
            presenter.Setup(p => p.Show()).Returns(model);
            presenter.Setup(p => p.Show(It.IsAny<Declaration>())).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRenamePresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}