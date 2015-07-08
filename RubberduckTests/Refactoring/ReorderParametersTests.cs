using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ReorderParametersTests : RefactoringTestBase
    {
        [TestMethod]
        public void ReorderParams_SwapPositions()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParams_WithOptionalParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParams_SwapPositions_UpdatesCallers()
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
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Bar()
 Foo ""Hello"", 10
End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderGetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg3 As Date, ByVal arg1 As Integer) As Boolean
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[2],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }

/*        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLetter()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date) As Boolean
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(Project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to remove
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, Module.Object.Lines());
        }*/

        private static Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>> SetupFactory(ReorderParametersModel model)
        {
            var presenter = new Mock<IReorderParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }
    }
}
