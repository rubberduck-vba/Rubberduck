using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ReorderParametersTests : RefactoringTestBase
    {
        private Mock<VBProject> _project;
        private Mock<VBComponent> _component;
        private Mock<CodeModule> _module;
        private List<Mock<CodeModule>> _modules;

        [TestCleanup]
        private void CleanUp()
        {
            _project = null;
            _component = null;
            _module = null;
        }

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
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
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
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderNamedParams()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg3:=6.1, arg1:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg3 As Double, ByVal arg2 As String)
End Sub

Public Sub Goo()
 Foo arg2:=""test44"", arg1:=3, arg3:=6.1
End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[0],
                model.Parameters[2],
                model.Parameters[1]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
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
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLetter()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLastParamFromSetter_NotAllowed()
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

            var model = new ReorderParametersModel(parseResult, qualifiedSelection);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderLastParamFromLetter_NotAllowed()
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

            var model = new ReorderParametersModel(parseResult, qualifiedSelection);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from letter
        }

        [TestMethod]
        public void ReorderParametersRefactoring_SignatureOnMultipleLines()
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
@"Private Sub Foo(ByVal arg3 As Date,                  ByVal arg2 As String,                  ByVal arg1 As Integer)


End Sub";   // note: IDE removes excess spaces

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[2],
                model.Parameters[1],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_CallOnMultipleLines()
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
@"Private Sub Foo(ByVal arg3 As Date, ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

 Foo arg3, arg2, arg1



End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[2],
                model.Parameters[1],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ClientReferencesAreNotUpdated_ParamArray()
        {
            //Input
            const string inputCode =
@"Sub Foo(ByVal arg1 As String, ParamArray arg2())
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Sub Foo(ByVal arg1 As String, ParamArray arg2())
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParams_MoveOptionalParamBeforeNonOptionalParamFails()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //set up model
            var model = new ReorderParametersModel(parseResult, qualifiedSelection);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[2],
                model.Parameters[0],
                model.Parameters[1]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParams_ReorderCallsWithoutOptionalParams()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg1, arg2
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
 Foo arg2, arg1
End Sub
";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderFirstParamFromGetterAndSetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Param(s) to reorder
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_ReorderFirstParamFromGetterAndLetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            //Arrange
            SetupProject(inputCode);
            var parseResult = new RubberduckParser().Parse(_project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            //Specify Params to reorder
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        #region setup
        private QualifiedSelection GetQualifiedSelection(Selection selection)
        {
            return new QualifiedSelection(new QualifiedModuleName(_component.Object), selection);
        }

        private static Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>> SetupFactory(ReorderParametersModel model)
        {
            var presenter = new Mock<IReorderParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>>();
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

            var components = MockFactory.CreateComponentsMock(new List<VBComponent>() { _component.Object });
            components.SetupGet(c => c.Parent).Returns(_project.Object);

            _project.SetupGet(p => p.VBComponents).Returns(components.Object);
            _component.SetupGet(c => c.Collection).Returns(components.Object);
        }

        private void SetupProject(params string[] inputCode)
        {
            var window = MockFactory.CreateWindowMock(string.Empty);
            var windows = new Mocks.MockWindowsCollection(window.Object);

            var vbe = MockFactory.CreateVbeMock(windows);

            foreach (var input in inputCode)
            {
                var codePane = MockFactory.CreateCodePaneMock(vbe, window);
                _modules.Add(MockFactory.CreateCodeModuleMock(input, codePane.Object));
            }

            _project = MockFactory.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);

            _component = MockFactory.CreateComponentMock("Module1", _module.Object, vbext_ComponentType.vbext_ct_StdModule);

            var components = MockFactory.CreateComponentsMock(new List<VBComponent>() { _component.Object });
            components.SetupGet(c => c.Parent).Returns(_project.Object);

            _project.SetupGet(p => p.VBComponents).Returns(components.Object);
            _component.SetupGet(c => c.Collection).Returns(components.Object);
        }

        #endregion
    }
}
