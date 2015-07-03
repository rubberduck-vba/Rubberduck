using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
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
            var window = Mocks.MockFactory.CreateWindowMock(string.Empty);
            var windows = new Mocks.MockWindowsCollection(window.Object);

            var mainWindow = new Mock<Window>();
            mainWindow.Setup(w => w.HWnd).Returns(0);

            var vbe = Mocks.MockFactory.CreateVbeMock(windows);
            vbe.SetupGet(v => v.MainWindow).Returns(mainWindow.Object);

            var codePane = new Mock<CodePane>();
            codePane.Setup(p => p.SetSelection(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>()));
            codePane.Setup(p => p.Show());
            codePane.SetupGet(p => p.VBE).Returns(vbe.Object);
            codePane.SetupGet(p => p.Window).Returns(window.Object);

            _module = Mocks.MockFactory.CreateCodeModuleMock(inputCode);
            _module.SetupGet(m => m.CodePane).Returns(codePane.Object);
            _module.Setup(m => m.ReplaceLine(It.IsAny<int>(), It.IsAny<string>()))
                .Callback<int, string>((i, s) => ReplaceModuleLine(_module, i, s));

            _project = Mocks.MockFactory.CreateProjectMock("VBAProject", vbext_ProjectProtection.vbext_pp_none);

            _component = Mocks.MockFactory.CreateComponentMock("Module1", _module.Object, vbext_ComponentType.vbext_ct_StdModule);

            var components = Mocks.MockFactory.CreateComponentsMock(new List<VBComponent>() {_component.Object});
            components.SetupGet(c => c.Parent).Returns(_project.Object);

            _project.SetupGet(p => p.VBComponents).Returns(components.Object);
            _component.SetupGet(c => c.Collection).Returns(components.Object);
        }

        private static void ReplaceModuleLine(Mock<CodeModule> module, int lineNumber, string newLine)
        {
            var lines = module.Object.Lines().Split(new[] {Environment.NewLine}, StringSplitOptions.None);

            lines[lineNumber - 1] = newLine;

            var newCode = String.Join(Environment.NewLine, lines);

            module.SetupGet(c => c.get_Lines(1, lines.Length)).Returns(newCode);
        }
        #endregion
    }
}
