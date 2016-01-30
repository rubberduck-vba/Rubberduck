using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class PromoteLocalToParameter
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
        public void PromoteLocalToParameterRefactoring_NoParamsInList()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal bar As Boolean)
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
            var refactoring = new IntroduceParameter(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_OneParamInList()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer)
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, ByVal bar As Boolean)
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
            var refactoring = new IntroduceParameter(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void PromoteLocalToParameterRefactoring_MultipleParamsOnMultipleLines()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean
End Sub";
            var selection = new Selection(3, 8, 3, 20); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer,                  ByRef baz As Date, ByVal bar As Boolean)
End Sub";   // note: the VBE removes extra spaces

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
            var refactoring = new IntroduceParameter(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
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
@"Private Sub Foo(ByVal buz As Integer,                  ByRef baz As Date, ByVal bar As Boolean)
    Dim bat As Date, bap As Integer
End Sub";   // note: the VBE removes extra spaces

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
            var refactoring = new IntroduceParameter(parser.State, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }
    }
}
