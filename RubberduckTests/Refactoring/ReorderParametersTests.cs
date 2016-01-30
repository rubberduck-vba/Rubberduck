using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;
using MessageBox = Rubberduck.UI.MessageBox;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class ReorderParametersTests : VbeTestBase
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParams_RefactorDeclaration()
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(model.TargetDeclaration);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParams_RefactorDeclaration_FailsInvalidTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);

            //assert
            try
            {
                refactoring.Refactor(
                    model.Declarations.FirstOrDefault(
                        i => i.DeclarationType == Rubberduck.Parsing.Symbols.DeclarationType.Module));
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Invalid declaration type", e.Message);
                return;
            }

            Assert.IsTrue(false);
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ReorderNamedParams_Function()
        {
            //Input
            const string inputCode =
@"Public Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public Function Foo(ByVal arg2 As String, ByVal arg1 As Integer) As Boolean
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
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ReorderNamedParams_WithOptionalParam()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Public Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
 Foo arg1:=3, arg2:=""test44""
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
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
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ReorderLastParamFromSetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ReorderLastParamFromLetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from letter
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
        {
            //Input
            const string inputCode =
@"Sub Foo(ByVal arg1 As String, ByVal arg2 As Date, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", arg, test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Sub Foo(ByVal arg2 As Date, ByVal arg1 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
 Foo arg, ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ClientReferencesAreUpdated_ParamArray_CallOnMultiplelines()
        {
            //Input
            const string inputCode =
@"Sub Foo(ByVal arg1 As String, ByVal arg2 As Date, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", _
        arg, _
        test1x, _
        test2x, _
        test3x, _
        test4x, _
        test5x, _
        test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Sub Foo(ByVal arg2 As Date, ByVal arg1 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
 Foo arg, ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParams_MoveOptionalParamBeforeNonOptionalParamFails()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[2],
                model.Parameters[0]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, MessageBoxIcon.Warning)).Returns(DialogResult.OK);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            var reorderedParams = new List<Parameter>()
            {
                model.Parameters[1],
                model.Parameters[0],
                model.Parameters[2]
            };

            model.Parameters = reorderedParams;

            var factory = SetupFactory(model);

            //act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
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
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to reorder
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
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
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null,
                parseResult.State, null);

            //act
            var refactoring = new ReorderParametersRefactoring(factory, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor();

            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_ParamsHaveDifferentNames()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer, ByVal v2 As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_ParamsHaveDifferentNames_TwoImplementations()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer, ByVal v2 As String)
End Sub";
            const string inputCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var module3 = project.Object.VBComponents.Item(2).CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_AcceptPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
@"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.Yes);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, messageBox.Object);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_ParamsSwapped_RejectPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.No);

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, messageBox.Object);
            Assert.IsNull(model.TargetDeclaration);
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_EventParamsSwapped()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_EventParamsSwapped_DifferentParamNames()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void ReorderParametersRefactoring_EventParamsSwapped_DifferentParamNames_TwoHandlers()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";
            const string inputCode3 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v1 As Integer, ByVal v2 As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode3 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .AddComponent("Class3", vbext_ComponentType.vbext_ct_ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var module3 = project.Object.VBComponents.Item(2).CodeModule;

            //Specify Params to remove
            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, null);
            model.Parameters.Reverse();

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new ReorderParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory), null);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void Presenter_AcceptDialog_ReordersProcedureWithTwoParameters()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, new MessageBox());
            model.Parameters.Reverse();

            var view = new Mock<IReorderParametersView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);
            view.Setup(v => v.Parameters).Returns(model.Parameters);

            var factory = new ReorderParametersPresenterFactory(editor.Object, view.Object, parseResult.State, null);

            var presenter = factory.Create();

            Assert.AreEqual(model.Parameters, presenter.Show().Parameters);
        }

        [TestMethod, Timeout(1000)]
        public void Presenter_CancelDialogCreatesNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new ReorderParametersModel(parseResult.State, qualifiedSelection, new MessageBox());

            var view = new Mock<IReorderParametersView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);
            view.Setup(v => v.Parameters).Returns(model.Parameters);

            var factory = new ReorderParametersPresenterFactory(editor.Object, view.Object, parseResult.State, null);
            var presenter = factory.Create();

            //Act
            var result = presenter.Show();

            //Assert
            Assert.IsNull(result);
        }

        [TestMethod, Timeout(1000)]
        public void Presenter_ParameterlessMemberCreatesNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null, parseResult.State, messageBox.Object);
            var presenter = factory.Create();

            //Act
            var result = presenter.Show();

            //Assert
            Assert.IsNull(result);
        }

        [TestMethod, Timeout(1000)]
        public void Presenter_SingleParameterMemberCreatesNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null, parseResult.State, messageBox.Object);
            var presenter = factory.Create();

            //Act
            var result = presenter.Show();

            //Assert
            Assert.IsNull(result);
        }

        [TestMethod, Timeout(1000)]
        public void Presenter_NullTargetCreatesNullModel()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 1, 1, 1); //startLine, startCol, endLine, endCol

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null, parseResult.State, null);

            var presenter = factory.Create();

            //Act
            var result = presenter.Show();

            //Assert
            Assert.IsNull(result);
        }

        [TestMethod, Timeout(1000)]
        public void Factory_NullSelectionCreatesNullPresenter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new ReorderParametersPresenterFactory(editor.Object, null, parseResult.State, null);

            //Act
            var result = factory.Create();

            //Assert
            Assert.IsNull(result);
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>> SetupFactory(ReorderParametersModel model)
        {
            var presenter = new Mock<IReorderParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IReorderParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}
