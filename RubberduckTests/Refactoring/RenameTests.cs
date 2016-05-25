using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RenameTests : VbeTestBase
    {
        [TestMethod]
        public void RenameRefactoring_RenameSub()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            var actual = module.Lines();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void RenameRefactoring_RenameVariable()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim val1 As Integer
End Sub";
            var selection = new Selection(2, 12, 2, 12);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim val2 As Integer
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameSub_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
    Foo
End Sub
";
            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode =
@"Private Sub Hoo()
End Sub

Private Sub Goo()
    Hoo
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameVariable_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim val1 As Integer
    val1 = val1 + 5
End Sub";
            var selection = new Selection(2, 12, 2, 12);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim val2 As Integer
    val2 = val2 + 5
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
    arg1 = ""test""
End Sub";
            var selection = new Selection(1, 25, 1, 25);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
    arg2 = ""test""
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFirstPropertyParameter_UpdatesAllRelatedParameters()
        {
            //Input
            const string inputCode =
@"Property Get Foo(ByVal index As Integer) As Variant
    Dim d As Integer
    d = index
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property

Property Set Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Integer
    d = index
End Property";
            var selection = new Selection(1, 28, 1, 28);

            //Expectation
            const string expectedCode =
@"Property Get Foo(ByVal renamed As Integer) As Variant
    Dim d As Integer
    d = renamed
End Property

Property Let Foo(ByVal renamed As Integer, ByVal value As Variant)
    Dim d As Integer
    d = renamed
End Property

Property Set Foo(ByVal renamed As Integer, ByVal value As Variant)
    Dim d As Integer
    d = renamed
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameLastPropertyParameter_UpdatesAllRelatedParameters()
        {
            //Input
            const string inputCode =
@"Property Get Foo(ByVal index As Integer) As Variant
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Variant
    d = value
End Property

Property Set Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Variant
    d = value
End Property";
            var selection = new Selection(4, 50, 4, 50);

            //Expectation
            const string expectedCode =
@"Property Get Foo(ByVal index As Integer) As Variant
End Property

Property Let Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property

Property Set Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameLastPropertyParameter_UpdatesRelatedParametersWithSameName()
        {
            //Input
            const string inputCode =
@"Property Get Foo(ByVal index As Integer) As Variant
End Property

Property Let Foo(ByVal index As Integer, ByVal value As Variant)
    Dim d As Variant
    d = value
End Property

Property Set Foo(ByVal index As Integer, ByVal fizz As Variant)
    Dim d As Variant
    d = fizz
End Property";
            var selection = new Selection(4, 50, 4, 50);

            //Expectation
            const string expectedCode =
@"Property Get Foo(ByVal index As Integer) As Variant
End Property

Property Let Foo(ByVal index As Integer, ByVal renamed As Variant)
    Dim d As Variant
    d = renamed
End Property

Property Set Foo(ByVal index As Integer, ByVal fizz As Variant)
    Dim d As Variant
    d = fizz
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameGetterAndSetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) 
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 25, 1, 25);

            //Expectation
            const string expectedCode =
@"Private Property Get Goo(ByVal arg1 As Integer) 
End Property

Private Property Set Goo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameGetterAndLetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo() 
End Property

Private Property Let Foo(ByVal arg1 As String) 
End Property";
            var selection = new Selection(1, 25, 1, 25);

            //Expectation
            const string expectedCode =
@"Private Property Get Goo() 
End Property

Private Property Let Goo(ByVal arg1 As String) 
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 21, 1, 21);

            //Expectation
            const string expectedCode =
@"Private Function Goo() As Boolean
    Goo = True
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFunction_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Foo = True
End Function

Private Sub Goo()
    Dim var1 As Boolean
    var1 = Foo()
End Sub
";
            var selection = new Selection(1, 21, 1, 21);

            //Expectation
            const string expectedCode =
@"Private Function Hoo() As Boolean
    Hoo = True
End Function

Private Sub Goo()
    Dim var1 As Boolean
    var1 = Hoo()
End Sub
";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RefactorWithDeclaration()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15);

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(model.Target);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterface()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(1, 22, 1, 22);

            //Expectation
            const string expectedCode1 =
@"Public Sub DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("IClass1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameEvent()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode1 =
@"Public Event Goo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Goo(ByVal i As Integer, ByVal s As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode1)
                .AddComponent("Class2", vbext_ComponentType.vbext_ct_ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_InterfaceRenamed_AcceptPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 27, 3, 27);

            //Expectation
            const string expectedCode1 =
@"Implements IClass1

Private Sub IClass1_DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";
            const string expectedCode2 =
@"Public Sub DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";

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
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, messageBox.Object) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, messageBox.Object, parser.State);
            refactoring.Refactor(model.Selection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_InterfaceRenamed_RejectPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 27);

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
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.No);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, messageBox.Object);
            Assert.AreEqual(null, model.Target);
        }

        [TestMethod]
        public void Rename_PresenterIsNull()
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
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            int startLine, startColumn, endLine, endColumn;

            var codePaneMock = new Mock<CodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(module);
            codePaneMock.Setup(c => c.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var factory = new RenamePresenterFactory(vbe.Object, null, parser.State, null);

            //act
            var refactoring = new RenameRefactoring(vbe.Object, factory, null, parser.State);
            refactoring.Refactor();

            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod]
        public void Presenter_TargetIsNull()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            int startLine, startColumn, endLine, endColumn;

            var codePaneMock = new Mock<CodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(module);
            codePaneMock.Setup(c => c.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var factory = new RenamePresenterFactory(vbe.Object, null, parser.State, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Factory_SelectionIsNull()
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
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            int startLine, startColumn, endLine, endColumn;

            var codePaneMock = new Mock<CodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(module);
            codePaneMock.Setup(c => c.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var factory = new RenamePresenterFactory(vbe.Object, null, parser.State, null);

            var presenter = factory.Create();
            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Factory_SelectionIsNotNull_Accept()
        {
            const string newName = "Goo";

            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, null) { NewName = newName };

            var view = new Mock<IRenameDialog>();
            view.Setup(v => v.NewName).Returns(newName);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var factory = new RenamePresenterFactory(vbe.Object, view.Object, parser.State, msgbox.Object);

            var presenter = factory.Create();
            Assert.AreEqual(model.NewName, presenter.Show().NewName);
        }

        [TestMethod]
        public void Factory_SelectionIsNotNull_Reject()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            int startLine, startColumn, endLine, endColumn;

            var codePaneMock = new Mock<CodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(module);
            codePaneMock.Setup(c => c.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var rubberduckCodePane = new Mock<ICodePaneWrapper>();
            rubberduckCodePane.Setup(r => r.CodeModule).Returns(module);
            rubberduckCodePane.Setup(r => r.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            rubberduckCodePane.Setup(r => r.Selection).Returns(selection);

            var codePaneFactoryMock = new Mock<ICodePaneWrapperFactory>();
            codePaneFactoryMock.Setup(c => c.Create(codePaneMock.Object)).Returns(rubberduckCodePane.Object);

            var view = new Mock<IRenameDialog>();
            view.Setup(v => v.NewName).Returns("Goo");
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);

            var factory = new RenamePresenterFactory(vbe.Object, view.Object, parser.State, null);

            var presenter = factory.Create();
            Assert.AreEqual(null, presenter.Show());
        }

        [TestMethod]
        public void Factory_SelectionIsNotNull_PassInTarget_Accept()
        {
            const string newName = "Goo";

            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, null) { NewName = newName };

            int startLine, startColumn, endLine, endColumn;

            var codePaneMock = new Mock<CodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(module);
            codePaneMock.Setup(c => c.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var rubberduckCodePane = new Mock<ICodePaneWrapper>();
            rubberduckCodePane.Setup(r => r.CodeModule).Returns(module);
            rubberduckCodePane.Setup(r => r.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            rubberduckCodePane.Setup(r => r.Selection).Returns(selection);

            var codePaneFactoryMock = new Mock<ICodePaneWrapperFactory>();
            codePaneFactoryMock.Setup(c => c.Create(codePaneMock.Object)).Returns(rubberduckCodePane.Object);

            var view = new Mock<IRenameDialog>();
            view.Setup(v => v.NewName).Returns(newName);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var factory = new RenamePresenterFactory(vbe.Object, view.Object, parser.State, null);

            var presenter = factory.Create();
            Assert.AreEqual(model.NewName, presenter.Show(model.Target).NewName);
        }

        [TestMethod]
        public void Factory_SelectionIsNotNull_PassInTarget_Reject()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25);

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, null) { NewName = "Goo" };

            int startLine, startColumn, endLine, endColumn;

            var codePaneMock = new Mock<CodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(module);
            codePaneMock.Setup(c => c.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var rubberduckCodePane = new Mock<ICodePaneWrapper>();
            rubberduckCodePane.Setup(r => r.CodeModule).Returns(module);
            rubberduckCodePane.Setup(r => r.GetSelection(out startLine, out startColumn, out endLine, out endColumn));
            rubberduckCodePane.Setup(r => r.Selection).Returns(selection);

            var codePaneFactoryMock = new Mock<ICodePaneWrapperFactory>();
            codePaneFactoryMock.Setup(c => c.Create(codePaneMock.Object)).Returns(rubberduckCodePane.Object);

            var view = new Mock<IRenameDialog>();
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.Cancel);

            var factory = new RenamePresenterFactory(vbe.Object, view.Object, parser.State, null);

            var presenter = factory.Create();
            Assert.AreEqual(null, presenter.Show(model.Target));
        }

        [TestMethod]
        public void RenameRefactoring_RenameProject()
        {
            const string oldName = "TestProject1";
            const string newName = "Renamed";

            //Arrange
            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder(oldName, vbext_ProjectProtection.vbext_pp_none)
                             .AddComponent("Module1", vbext_ComponentType.vbext_ct_StdModule, string.Empty)
                             .MockVbeBuilder()
                             .Build();
            
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());
            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, default(QualifiedSelection), msgbox.Object) { NewName = newName };
            model.Target = model.Declarations.First(i => i.DeclarationType == DeclarationType.Project && !i.IsBuiltIn);

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(model.Target);

            //Assert
            Assert.AreEqual(newName, vbe.Object.VBProjects.Item(0).Name);
        }

        [TestMethod]
        public void RenameRefactoring_RenameSub_ConflictingNames_Reject()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim Goo As Integer
End Sub";
            var selection = new Selection(1, 14, 1, 14);

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim Goo As Integer
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.No);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, messageBox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameSub_ConflictingNames_Accept()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim Goo As Integer
End Sub";
            var selection = new Selection(1, 14, 1, 14);

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
    Dim Goo As Integer
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.Yes);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, messageBox.Object, parser.State);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameCodeModule()
        {
            const string newName = "RenameModule";

            //Input
            const string inputCode =
@"Private Sub Foo(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 27, 3, 27);

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Class1", vbext_ComponentType.vbext_ct_ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents.Item(0);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parser.State, qualifiedSelection, msgbox.Object) { NewName = newName };
            model.Target = model.Declarations.FirstOrDefault(i => i.DeclarationType == DeclarationType.ClassModule && i.IdentifierName == "Class1");

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(vbe.Object, factory.Object, msgbox.Object, parser.State);
            refactoring.Refactor(model.Target);

            //Assert
            Assert.AreSame(newName, component.CodeModule.Name);
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IRenamePresenter>> SetupFactory(RenameModel model)
        {
            var presenter = new Mock<IRenamePresenter>();
            presenter.Setup(p => p.Show()).Returns(model);
            presenter.Setup(p => p.Show(It.IsAny<Declaration>())).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRenamePresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}
