using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ReorderParametersTests
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
            var refactoring = new ReorderParametersRefactoring(factory.Object);
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
            var refactoring = new ReorderParametersRefactoring(factory.Object);
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
            var refactoring = new ReorderParametersRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, _module.Object.Lines());
        }

        private static Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>> SetupFactory(ReorderParametersModel model)
        {
            var presenter = new Mock<IReorderParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        private QualifiedSelection GetQualifiedSelection(Selection selection)
        {
            return new QualifiedSelection(new QualifiedModuleName(_component.Object), selection);
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
    }
}
