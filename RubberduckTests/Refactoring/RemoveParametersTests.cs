using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RemoveParametersTests : RefactoringTestBase
    {
        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveBothParams()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo( )
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters.ForEach(arg => arg.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveOnlyParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters.ForEach(arg => arg.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveFirstParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo( ByVal arg2 As String)
End Sub"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveSecondParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer )
End Sub"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveFromFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer ) As Boolean
End Function"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveFromGetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo() As Boolean
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters.ForEach(p => p.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        //bug: We shouldn't allow the only param in a setter to be removed, it will break the VBA code.
        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveFromSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Set Foo()
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters.ForEach(p => p.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        //note: removing other params from setters is fine (In fact, we may want to create an inspection for this).
        [TestMethod]
        public void RemoveParamatersRefactoring_RemoveSecondParamFromSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Set Foo(ByVal arg1 As Integer )
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParamatersRefactoring_ClientReferencesAreUpdated()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo 10, ""Hello""
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer )
End Sub

Private Sub Bar()
 Foo 10 
End Sub
"; //note: The IDE strips out the extra whitespace, you can't see it but there's a space after "Foo 10 "

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        private static Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>> SetupFactory(RemoveParametersModel model)
        {
            var presenter = new Mock<IRemoveParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }
    }
}
