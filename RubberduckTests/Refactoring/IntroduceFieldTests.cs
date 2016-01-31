using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class IntroduceFieldTests
    {
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(0, 1);

        void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Ready)
            {
                _semaphore.Release();
            }
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_NoFieldsInClass_Sub()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Sub Foo()
    
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_NoFieldsInList_Function()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Dim bar As Boolean
    Foo = True
End Function";
            var selection = new Selection(2, 10, 2, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Function Foo() As Boolean
    
    Foo = True
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_OneFieldInList()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Private Sub Foo(ByVal buz As Integer)
    Dim bar As Boolean
End Sub";
            var selection = new Selection(3, 10, 3, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public fizz As Integer
Private bar As Boolean
Private Sub Foo(ByVal buz As Integer)
    
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_MultipleFieldsOnMultipleLines()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Public buzz As Integer
Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean
End Sub";
            var selection = new Selection(5, 8, 5, 20); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public fizz As Integer
Public buzz As Integer
Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_MultipleVariablesInStatement_MoveFirst()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date, _
        bap As Integer
End Sub";
            var selection = new Selection(3, 10, 3, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim _
        bat As Date, _
        bap As Integer
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_MultipleVariablesInStatement_MoveSecond()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date, _
        bap As Integer
End Sub";
            var selection = new Selection(4, 10, 4, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bat As Date
Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
         _
        bap As Integer
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_MultipleVariablesInStatement_MoveLast()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date, _
        bap As Integer
End Sub";
            var selection = new Selection(5, 10, 5, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bap As Integer
Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date
        
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_MultipleVariablesInStatement_OnOneLine_MoveFirst()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, bat As Date, bap As Integer
End Sub";
            var selection = new Selection(3, 10, 3, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bat As Date, bap As Integer
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_DisplaysInvalidSelectionAndDoesNothingForField()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean

Private Sub Foo()
End Sub";
            var selection = new Selection(1, 14, 1, 14); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_DisplaysInvalidSelectionAndDoesNothingForInvalidSelection()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean

Private Sub Foo()
End Sub";
            var selection = new Selection(3, 16, 3, 16); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_PassInTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Sub Foo()
    
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(parser.State.AllUserDeclarations.FindVariable(qualifiedSelection));

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_PassInTarget_Nonvariable()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            //Act
            var refactoring = new IntroduceField(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);

            //Assert
            try
            {
                refactoring.Refactor(parser.State.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));
                messageBox.Verify(m =>
                        m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                            It.IsAny<MessageBoxIcon>()), Times.Once);
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Invalid declaration type", e.Message);
                Assert.AreEqual(inputCode, module.Lines());
                return;
            }

            Assert.Fail();
        }
    }
}
