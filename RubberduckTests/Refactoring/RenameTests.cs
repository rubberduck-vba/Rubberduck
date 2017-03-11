using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RenameTests
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void RenameRefactoring_RenameMulitlinedParameter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String, _
        ByVal arg3 As String)
End Sub";
            var selection = new Selection(2, 15, 2, 15);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg1 As String, _
        ByVal arg2 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(model.Target);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, messageBox.Object) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(model.Selection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(project.Object.VBComponents[0]), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.No);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, messageBox.Object);
            Assert.AreEqual(null, model.Target);
        }

        [TestMethod]
        public void Rename_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var codePaneMock = new Mock<ICodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(component.CodeModule);
            codePaneMock.Setup(c => c.Selection);
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var vbeWrapper = vbe.Object;
            var factory = new RenamePresenterFactory(vbeWrapper, null, state, null);

            var refactoring = new RenameRefactoring(vbeWrapper, factory, null, state);
            refactoring.Refactor();

            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void Presenter_TargetIsNull()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var codePaneMock = new Mock<ICodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(component.CodeModule);
            codePaneMock.Setup(c => c.Selection);
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var vbeWrapper = vbe.Object;
            var factory = new RenamePresenterFactory(vbeWrapper, null, state, null);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var codePaneMock = new Mock<ICodePane>();
            codePaneMock.Setup(c => c.CodeModule).Returns(component.CodeModule);
            codePaneMock.Setup(c => c.Selection);
            vbe.Setup(v => v.ActiveCodePane).Returns(codePaneMock.Object);

            var vbeWrapper = vbe.Object;
            var factory = new RenamePresenterFactory(vbeWrapper, null, state, null);

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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, null) { NewName = newName };

            var view = new Mock<IRenameDialog>();
            view.Setup(v => v.NewName).Returns(newName);
            view.Setup(v => v.ShowDialog()).Returns(DialogResult.OK);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var factory = new RenamePresenterFactory(vbeWrapper, view.Object, state, msgbox.Object);

            var presenter = factory.Create();
            Assert.AreEqual(model.NewName, presenter.Show().NewName);
        }

        [TestMethod]
        public void RenameRefactoring_RenameProject()
        {
            const string oldName = "TestProject1";
            const string newName = "Renamed";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder(oldName, ProjectProtection.Unprotected)
                             .AddComponent("Module1", ComponentType.StandardModule, string.Empty)
                             .MockVbeBuilder()
                             .Build();


            var state = MockParser.CreateAndParse(vbe.Object);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, default(QualifiedSelection), msgbox.Object) { NewName = newName };
            model.Target = model.Declarations.First(i => i.DeclarationType == DeclarationType.Project && !i.IsBuiltIn);

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(model.Target);

            Assert.AreEqual(newName, vbe.Object.VBProjects[0].Name);
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.No);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.Yes);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, "Class1", ComponentType.ClassModule, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection, msgbox.Object) { NewName = newName };
            model.Target = model.Declarations.FirstOrDefault(i => i.DeclarationType == DeclarationType.ClassModule && i.IdentifierName == "Class1");

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(model.Target);

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
