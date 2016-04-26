using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;
using MessageBox = Rubberduck.UI.MessageBox;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RemoveParametersTests : VbeTestBase
    {
        [TestMethod]
        public void RemoveParametersRefactoring_RemoveBothParams()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo( )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters.ForEach(arg => arg.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveOnlyParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters.ForEach(arg => arg.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFirstParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo( ByVal arg2 As String)
End Sub"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveSecondParam()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer )
End Sub"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveNamedParam()
        {
            //Input
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg3:=6.1, arg1:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String )
End Sub

Public Sub Goo()
    Foo arg2:=""test44"",  arg1:=3
End Sub
"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[2].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveLastFromFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Function Foo(ByVal arg1 As Integer ) As Boolean
End Function"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveAllFromFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
End Function";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Function Foo( ) As Boolean
End Function"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters.ForEach(p => p.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Function Foo( ) As Boolean
End Function

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo  
End Sub
"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters.ForEach(p => p.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFromGetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Get Foo() As Boolean
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters.ForEach(p => p.IsRemoved = true);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_QuickFix()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 38, 1, 38);

            //Expectation
            const string expectedCode =
@"Private Property Set Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.QuickFix(parser.State, qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveFirstParamFromSetter()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Set Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo( ByVal arg2 As String)
End Sub

Private Sub Bar()
    Foo  ""Hello""
End Sub
"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer )
End Sub

Private Sub Bar()
    Foo 10 
End Sub
"; //note: The IDE strips out the extra whitespace, you can't see it but there's a space after "Foo 10 "

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Sub Foo(ByVal arg1 As String )
End Sub

Public Sub Goo(ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test""      
End Sub
"; //note: The IDE strips out the extra whitespace, you can't see it but there are several spaces after " ParamArrayTest ""test""      "

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveLastParamFromSetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from setter
        }

        [TestMethod]
        public void RemoveParametersRefactoring_RemoveLastParamFromLetter_NotAllowed()
        {
            //Input
            const string inputCode =
@"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);

            // Assert
            Assert.AreEqual(1, model.Parameters.Count); // doesn't allow removing last param from letter
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Get Foo()
End Property

Private Property Set Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Property Get Foo()
End Property

Private Property Let Foo( ByVal arg2 As String)
End Property"; //note: The IDE strips out the extra whitespace

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo( Optional ByVal arg2 As String)
End Sub

Private Sub Goo(ByVal arg1 As Integer)
    Foo 
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(                  ByVal arg2 As String,                  ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines_RemoveSecond()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer,                                    ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_SignatureOnMultipleLines_RemoveLast()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As Integer,                  ByVal arg2 As String                  )
End Sub";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[2].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_PassTargetIn()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, _
                  ByVal arg2 As String, _
                  ByVal arg3 As Date)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(                  ByVal arg2 As String,                  ByVal arg3 As Date)
End Sub";   // note: VBE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(model.TargetDeclaration);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
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
            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode =
@"Private Sub Foo( ByVal arg2 As String, ByVal arg3 As Date)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo  arg2, arg3

End Sub
";   // note: IDE removes excess spaces

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[0].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal a As Integer )
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved_ImplementationParamsHaveDifferentNames()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer, ByVal v2 As String)
End Sub";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal a As Integer )
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved_ImplementationParamsHaveDifferentNames_TwoImplementations()
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

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode1 =
@"Public Sub DoSomething(ByVal a As Integer )
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v1 As Integer )
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal i As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var module3 = project.Object.VBComponents.Item(2).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastEventParamRemoved()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg1 As Integer )";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg1 As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void ReorderParametersRefactoring_LastEventParamRemoved_EventImplementationSelected()
        {
            //Input
            const string inputCode1 =
@"Private WithEvents abc As Class2

Private Sub abc_Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string inputCode2 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            var selection = new Selection(3, 15, 3, 15);

            //Expectation
            const string expectedCode1 =
@"Private WithEvents abc As Class2

Private Sub abc_Foo(ByVal arg1 As Integer )
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
@"Public Event Foo(ByVal arg1 As Integer )";

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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters.Last().IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastEventParamRemoved_ParamsHaveDifferentNames()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";

            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg1 As Integer )";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastEventParamRemoved_ParamsHaveDifferentNames_TwoHandlers()
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

            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode1 =
@"Public Event Foo(ByVal arg1 As Integer )";

            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer )
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v1 As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var module3 = project.Object.VBComponents.Item(2).CodeModule;

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
            Assert.AreEqual(expectedCode3, module3.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastInterfaceParamsRemoved_AcceptPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 23);

            //Expectation
            const string expectedCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer )
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
@"Public Sub DoSomething(ByVal a As Integer )
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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.Yes);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, messageBox.Object);
            model.Parameters[1].IsRemoved = true;

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RemoveParametersRefactoring_LastInterfaceParamRemoved_RejectPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 23);

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
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.No);

            //Specify Params to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, messageBox.Object);
            Assert.IsNull(model.TargetDeclaration);
        }

        [TestMethod]
        public void RemoveParams_RefactorDeclaration_FailsInvalidTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //set up model
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);

            var factory = SetupFactory(model);

            //act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));

            //assert
            try
            {
                refactoring.Refactor(
                    model.Declarations.FirstOrDefault(
                        i => i.DeclarationType == Rubberduck.Parsing.Symbols.DeclarationType.ProceduralModule));
            }
            catch (ArgumentException e)
            {
                Assert.AreEqual("Invalid declaration type", e.Message);
                return;
            }

            Assert.Fail();
        }

        [TestMethod]
        public void RemoveParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new RemoveParametersPresenterFactory(editor.Object, null,
                parser.State, null);

            //act
            var refactoring = new RemoveParametersRefactoring(factory, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor();

            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod]
        public void RemoveParams_ModelIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var module = component.CodeModule;
            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            //Specify Param(s) to remove
            var model = new RemoveParametersModel(parser.State, qualifiedSelection, null);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RemoveParametersRefactoring(factory.Object, new ActiveCodePaneEditor(vbe.Object, codePaneFactory));
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod]
        public void Presenter_Accept_ReturnsModelWithParametersChanged()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new RemoveParametersModel(parser.State, qualifiedSelection, new MessageBox());
            model.Parameters[1].IsRemoved = true;

            var view = new Mock<IRemoveParametersView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);
            view.Setup(v => v.Parameters).Returns(model.Parameters);

            var factory = new RemoveParametersPresenterFactory(editor.Object, view.Object, parser.State, null);

            var presenter = factory.Create();

            Assert.AreEqual(model.Parameters, presenter.Show().Parameters);
        }

        [TestMethod]
        public void Presenter_Reject_ReturnsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new RemoveParametersModel(parser.State, qualifiedSelection, new MessageBox());
            model.Parameters[1].IsRemoved = true;

            var view = new Mock<IRemoveParametersView>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);
            view.Setup(v => v.Parameters).Returns(model.Parameters);

            var factory = new RemoveParametersPresenterFactory(editor.Object, view.Object, parser.State, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_Accept_AutoMarksSingleParamAsRemoved()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As Integer)
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var model = new RemoveParametersModel(parser.State, qualifiedSelection, new MessageBox());
            model.Parameters[0].IsRemoved = true;

            var factory = new RemoveParametersPresenterFactory(editor.Object, null, parser.State, null);

            var presenter = factory.Create();

            Assert.IsTrue(model.Parameters[0].Declaration.Equals(presenter.Show().Parameters[0].Declaration));
        }

        [TestMethod]
        public void Presenter_ParameterlessTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectBuilder.AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var codePane = project.Object.VBComponents.Item(0).CodeModule.CodePane;
            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var factory = new RemoveParametersPresenterFactory(editor.Object, null, parser.State, messageBox.Object);
            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Presenter_NullTargetReturnsNullModel()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = Selection.Home;

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectBuilder.AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var codePaneFactory = new CodePaneWrapperFactory();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var codePane = project.Object.VBComponents.Item(0).CodeModule.CodePane;
            var ext = codePaneFactory.Create(codePane);
            ext.Selection = selection;

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), selection);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns(qualifiedSelection);

            var factory = new RemoveParametersPresenterFactory(editor.Object, null, parser.State, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Factory_NullSelectionNullReturnsNullPresenter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            projectBuilder.AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, inputCode);
            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new RemoveParametersPresenterFactory(editor.Object, null, parser.State, null);

            Assert.AreEqual(null, factory.Create());
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>> SetupFactory(RemoveParametersModel model)
        {
            var presenter = new Mock<IRemoveParametersPresenter>();
            presenter.Setup(p => p.Show()).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRemoveParametersPresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}
