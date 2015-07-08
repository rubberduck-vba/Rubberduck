using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RemoveParametersTests
    {
        private Mock<VBProject> _project;
        private Mock<VBComponent> _component;
        private Mock<CodeModule> _module;

        [TestCleanup]
        private void CleanUp()
        {
            _project = null;
            _component = null;
            _module = null;
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveBothParams()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveOnlyParam()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFirstParam()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveSecondParam()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveLastFromFunction()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveAllFromFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Foo( ) As Boolean
End Function"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveAllFromFunction_UpdateCallReferences()
        {
            //Input
            const string inputCode =
@"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg1, arg2
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Foo( ) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
 Foo  
End Sub
"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFromGetter()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        //note: removing other params from setters is fine (In fact, we may want to create an inspection for this).
        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFirstParamFromSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Set Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_FirstParam()
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
@"Private Sub Foo( ByVal arg2 As String)
End Sub

Private Sub Bar()
 Foo  ""Hello""
End Sub
"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_LastParam()
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
        {
            //Input
            const string inputCode =
@"Sub ParamArrayTest(ByVal Hihi As String, ParamArray Hoho())
End Sub

Public Sub Haha(ByVal test1x As Integer, _
                ByVal test2x As Integer, _
                ByVal test3x As Integer, _
                ByVal test4x As Integer, _
                ByVal test5x As Integer, _
                ByVal test6x As Integer)
               
    ParamArrayTest ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Sub ParamArrayTest(ByVal Hihi As String )
End Sub

Public Sub Haha(ByVal test1x As Integer, _
                ByVal test2x As Integer, _
                ByVal test3x As Integer, _
                ByVal test4x As Integer, _
                ByVal test5x As Integer, _
                ByVal test6x As Integer)
               
 ParamArrayTest ""test""      
End Sub
"; //note: The IDE strips out the extra whitespace, you can't see it but there are several spaces after " ParamArrayTest ""test""      "

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveLastParamFromSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RemoveParametersModel(parseResult, qualifiedSelection);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveLastParamFromLetter()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RemoveParametersModel(parseResult, qualifiedSelection);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndSetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) 
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo()
End Property

Private Property Set Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFirstParamFromGetterAndLetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) 
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo()
End Property

Private Property Let Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_SignatureContainsOptionalParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo arg1
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo( Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
 Foo 
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[0].IsRemoved  = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(                  ByVal arg2 As String,                  ByVal arg3 As Date)


End Sub";   // note: IDE removes excess spaces

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_CallOnMultipleLines()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg1, _
        arg2, _
        arg3

End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo( ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

 Foo  arg2, arg3



End Sub
";   // note: IDE removes excess spaces

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parseResult, qualifiedSelection);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        #region setup
        private QualifiedSelection GetQualifiedSelection(Selection selection)
        {
            return new QualifiedSelection(new QualifiedModuleName(_component.Object), selection);
        }

        private static Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>> SetupFactory(RemoveParametersModel model)
        {
            var presenter = new Mock<IRemoveParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        private void SetupProject(string inputCode)
        {
            var window = MockFactory.CreateWindowMock(string.Empty);
            var windows = new Mocks.MockWindowsCollection(window.Object);

            var vbe = MockFactory.CreateVbeMock(windows);

            var codePane = MockFactory.CreateCodePaneMock(vbe, window);

            _module = MockFactory.CreateCodeModuleMock(inputCode, codePane.Object);
           
            _project = MockFactory.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);

            _component = MockFactory.CreateComponentMock("Module1", _module.Object, vbext_ComponentType.vbext_ct_StdModule);

            var components = MockFactory.CreateComponentsMock(new List<VBComponent>() {_component.Object});
            components.SetupGet(c => c.Parent).Returns(_project.Object);

            _project.SetupGet(p => p.VBComponents).Returns(components.Object);
            _component.SetupGet(c => c.Collection).Returns(components.Object);
        }

        #endregion
    }
}
