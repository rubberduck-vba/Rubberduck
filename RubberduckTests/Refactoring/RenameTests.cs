using System;
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
using System.Collections.Generic;
using Rubberduck.Parsing.VBA;

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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "val2" };

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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter_DoesNotAlterPrecompilerDirectives()
        {
            //Input
            const string inputCode =
@"#Const Bar = 42

#If False Then
Private Sub Goo(ByVal arg1 As String)
#ElseIf True Then
Private Sub Foo(ByVal arg1 As String)
#Else
Private Sub Foo(ByVal arg1 As String, arg2 As String)
#End If
End Sub";
            var selection = new Selection(6, 25, 6, 25);

            //Expectation
            const string expectedCode =
@"#Const Bar = 42

#If False Then
Private Sub Goo(ByVal arg1 As String)
#ElseIf True Then
Private Sub Foo(ByVal arg2 As String)
#Else
Private Sub Foo(ByVal arg1 As String, arg2 As String)
#End If
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "renamed" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var msgbox = new Mock<IMessageBox>();
            msgbox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(model.Target);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter1 = state.GetRewriter(module1.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = state.GetRewriter(module2.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());

            msgbox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterfaceReferences()
        {
            const string inputCode1 =
@"Public Sub DoSomething()
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub";
            const string inputCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoSomething
End Sub"
;
            const string expectedCode1 =
@"Public Sub DoNothing()
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub";

            const string expectedCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoNothing
End Sub"
;
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "DoSomething",
                SelectionLineIdentifier = "Sub DoSomething",
                SelectionModuleName = "IClass1",
                NewName = "DoNothing"
            };

            var secondClassName = "Class1";
            var thirdClassName = "Class3";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode1, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClassName, inputCode2, ComponentType.ClassModule);
            AddTestComponent(tdo, thirdClassName, inputCode3, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());

            var rewriter3 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, thirdClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode3, rewriter3.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterfaceFromImplementation()
        {
            const string inputCode1 =
@"Public Sub DoSomething()
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub";
            const string inputCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoSomething
End Sub"
;

            const string expectedCode1 =
@"Public Sub DoNothing()
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub";

            const string expectedCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoNothing
End Sub"
;
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "DoSomething",
                SelectionLineIdentifier = "IClass1_DoSomething(",
                SelectionModuleName = "Class1",
                NewName = "DoNothing"
            };

            var secondClassName = "IClass1";
            var thirdClassName = "Class3";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode2, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClassName, inputCode1, ComponentType.ClassModule);
            AddTestComponent(tdo, thirdClassName, inputCode3, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.Declaration);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter2.GetText());

            var rewriter3 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, thirdClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode3, rewriter3.GetText());

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Once);
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterfaceNoImplementers()
        {
            const string inputCode1 =
@"Public Sub DoSomething()
End Sub";

            const string expectedCode1 =
@"Public Sub DoNothing()
End Sub";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "DoSomething",
                SelectionLineIdentifier = "Sub DoSomething(",
                SelectionModuleName = "IClass1",
                NewName = "DoNothing"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode1, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.Declaration);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterfaceFromReference()
        {
            const string inputCode1 =
@"Public Sub DoSomething()
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub";
            const string inputCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoSomething
End Sub"
;

            const string expectedCode1 =
@"Public Sub DoNothing()
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub";

            const string expectedCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class1
    Dim c2 As IClass1
    Set c1 = new Class1
    Set c2 = c1
    c1.DoNothing
End Sub"
;
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "DoSomething",
                SelectionLineIdentifier = "c1.DoSomething",
                SelectionModuleName = "Class3",
                NewName = "DoNothing"
            };

            var secondClassName = "Class1";
            var thirdClassName = "IClass1";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode3, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClassName, inputCode2, ComponentType.ClassModule);
            AddTestComponent(tdo, thirdClassName, inputCode1, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode3, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());

            var rewriter3 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, thirdClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter3.GetText());
            
        }

        [TestMethod]
        public void RenameRefactoring_RenameControl()
        {
            const string inputCode1 =
@"
Private Sub cmdBtn1_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBtn1.Caption = ""Click This""
End Sub
";

            const string expectedCode1 =
@"
Private Sub cmdBigButton_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBigButton.Caption = ""Click This""
End Sub
";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "cmdBtn1_Click",
                SelectionLineIdentifier = "Private Sub cmdBtn1_Click()",
                SelectionModuleName = "UserForm1",
                NewName = "cmdBigButton"
            };

            CreateMockVBEForControlsTest(tdo, inputCode1, "cmdBtn1");

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameControlRenameInReference()
        {
            const string inputCode1 =
@"
Private Sub cmdBtn1_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBtn1.Caption = ""Click This""
End Sub
";

            const string expectedCode1 = 

@"
Private Sub cmdBigButton_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBigButton.Caption = ""Click This""
End Sub
";

            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "cmdBtn1",
                SelectionLineIdentifier = "cmdBtn1.Caption =",
                SelectionModuleName = "UserForm1",
                NewName = "cmdBigButton"
            };

            CreateMockVBEForControlsTest(tdo, inputCode1, "cmdBtn1");

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [TestMethod]
        public void RenameRefactoring_RenameControlFromEventReference()
        {
            const string inputCode1 =
@"
Private Sub cmdBtn1_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBtn1.Caption = ""Click This""
End Sub
";

            const string expectedCode1 =
@"
Private Sub cmdBigButton_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBigButton_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBigButton.Caption = ""Click This""
End Sub
";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "cmdBtn1_Click",
                SelectionLineIdentifier = "cmdBtn1_Click 'bad idea",
                SelectionModuleName = "UserForm1",
                NewName = "cmdBigButton"
            };

            CreateMockVBEForControlsTest(tdo, inputCode1, "cmdBtn1");

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [TestMethod]
        public void RenameRefactoring_RenameControlFromChangeEventHandler()
        {
            const string inputCode1 =
@"
Private Sub cmdBtn1_Click()

End Sub

Private Sub tbEnterName_Change()
    cmdBtn1_Click 'bad idea, but someone will do it
End Sub

Private Sub UserForm_Click()
    cmdBtn1.Caption = ""Click This""
End Sub
";

            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "cmdBtn1_Click",
                SelectionLineIdentifier = "Private Sub cmdBtn1_Click()",
                SelectionModuleName = "UserForm1",
                NewName = "cmdBtn1_ClickAgain"
            };

            CreateMockVBEForControlsTest(tdo, inputCode1, "cmdBtn1");

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.OK, It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        private void CreateMockVBEForControlsTest(RenameTestsDataObject tdo, string inputCode, string controlName)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder(tdo.ProjectName, ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder(tdo.SelectionModuleName, inputCode).AddControl(controlName).Build();
            project.AddComponent(form);
            builder.AddProject(project.Build());
            tdo.VBE = builder.Build().Object;
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterfaceReferencesWithinScope()
        {
            const string inputCode1 =
@"Public Sub DoSomething()
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething()
End Sub";
            const string inputCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoSomething
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class2
    Dim c2 As IClass1
    Set c1 = new Class2
    Set c2 = c1
    c1.DoSomething  'This is left alone because it is a member of Class2, not the interface
    c2.DoSomething
End Sub"
;

            const string expectedCode1 =
@"Public Sub DoNothing()
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoNothing()
End Sub";

            const string expectedCode3 =
@"Private Sub RefTheInterface()
    Dim c1 As Class1
    Set c1 = new IClass1
    c1.DoNothing
End Sub

Private Sub RefTheInterface2()
    Dim c1 As Class2
    Dim c2 As IClass1
    Set c1 = new Class2
    Set c2 = c1
    c1.DoSomething  'This is left alone because it is a member of Class2, not the interface
    c2.DoNothing
End Sub"
;
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "DoSomething",
                SelectionLineIdentifier = "Sub DoSomething",
                SelectionModuleName = "IClass1",
                NewName = "DoNothing"
            };

            var secondClassName = "Class1";
            var thirdClassName = "Class3";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode1, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClassName, inputCode2, ComponentType.ClassModule);
            AddTestComponent(tdo, "Class3", inputCode3, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());

            var rewriter3 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, thirdClassName).CodeModule.Parent);
            Assert.AreEqual(expectedCode3, rewriter3.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameEventWithReferences()
        {
            const string inputCode1 =
@"
Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent MyEvent(1234, Cancel)
End Sub
";
            const string inputCode2 =
@"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_MyEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub
";

            const string expectedCode1 =
@"
Public Event YourEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent YourEvent(1234, Cancel)
End Sub
";
            const string expectedCode2 =
@"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_YourEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub
";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "MyEvent",
                SelectionLineIdentifier = "Event MyEvent",
                SelectionModuleName = "CEventClass",
                NewName = "YourEvent"
            };

            var secondClass = "Class2";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode1, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClass, inputCode2, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClass).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameModuleFromReference()
        {
            const string inputCode1 =
@"
Sub Foo()
End Sub
";
            const string inputCode2 =
@"
Sub Foo2()
    Dim c1 As CTestClass
    Set c1 = new CTestClass
    c1.Foo
End Sub
";

            const string expectedCode1 =
@"
Sub Foo()
End Sub
";
            const string expectedCode2 =
@"
Sub Foo2()
    Dim c1 As CMyTestClass
    Set c1 = new CMyTestClass
    c1.Foo
End Sub
";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "CTestClass",
                SelectionLineIdentifier = "c1 As CTestClass",
                SelectionModuleName = "Class2",
                NewName = "CMyTestClass"
            };

            var secondClass = "CTestClass";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode2, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClass, inputCode1, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClass).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter2.GetText());

            var component = RetrieveComponent(tdo, secondClass);
            Assert.AreSame(tdo.NewName, component.CodeModule.Name);
        }

        [TestMethod]
        public void RenameRefactoring_RenameEventFromUsage()
        {
            const string inputCode1 =
@"
Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent MyEvent(1234, Cancel)
End Sub
";
            const string inputCode2 =
@"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_MyEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub
";

            const string expectedCode1 =
@"
Public Event YourEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent YourEvent(1234, Cancel)
End Sub
";
            const string expectedCode2 =
@"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_YourEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub
";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "MyEvent",
                SelectionLineIdentifier = "RaiseEvent MyEvent",
                SelectionModuleName = "CEventClass",
                NewName = "YourEvent"
            };

            var secondClass = "Class2";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode1, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClass, inputCode2, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClass).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameEventFromImplementer()
        {
            //CEventClass
            const string inputCode1 =
@"
Public Event MyEvent(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent MyEvent(1234, Cancel)
End Sub
";
            //Class2
            const string inputCode2 =
@"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_MyEvent(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub
";

            const string expectedCode1 =
@"
Public Event YourEvent_withUnderscore(IDNumber As Long, ByRef Cancel As Boolean)

Sub AAA()
    Dim Cancel As Boolean
    Cancel = False
    RaiseEvent YourEvent_withUnderscore(1234, Cancel)
End Sub
";
            const string expectedCode2 =
@"
Private WithEvents XLEvents As CEventClass

Private Sub Class_Initialize()
    Set XLEvents = New CEventClass
End Sub

Private Sub XLEvents_YourEvent_withUnderscore(IDNumber As Long, Cancel As Boolean)
    Cancel = True
End Sub
";
            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "MyEvent",
                SelectionLineIdentifier = "Private Sub XLEvents_MyEvent",
                SelectionModuleName = "Class2",
                NewName = "YourEvent_withUnderscore"
            };

            var secondClass = "CEventClass";
            AddTestComponent(tdo, tdo.SelectionModuleName, inputCode2, ComponentType.ClassModule);
            AddTestComponent(tdo, secondClass, inputCode1, ComponentType.ClassModule);

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);

            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode2, rewriter1.GetText());

            var rewriter2 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, secondClass).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter2.GetText());
        }

        [TestMethod]
        public void RenameRefactoring_RenameControlEventName_AcceptPrompt()
        {
            const string inputCode1 =
@"
Private Sub cmdBtn1_Click()

End Sub
";

            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "cmdBtn1_Click",
                SelectionLineIdentifier = "Private Sub cmdBtn1_Click()",
                SelectionModuleName = "UserForm1",
                NewName = "cmdBtn2"
            };

            CreateMockVBEForControlsTest(tdo, inputCode1, "cmdBtn1");

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.None);
            tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
        }

        [TestMethod]
        public void RenameRefactoring_CheckAllRefactorCallPaths()
        {
            const string inputCode1 =
@"
Private Sub Foo()
End Sub
";
            const string expectedCode1 =
@"
Private Sub Goo()
End Sub
";
            RefactorParams[] refactorParams = { RefactorParams.None, RefactorParams.QualifiedSelection, RefactorParams.Declaration };

            foreach ( var param in refactorParams)
            {
                var tdo = new RenameTestsDataObject
                {
                    SelectionTarget = "Foo",
                    SelectionLineIdentifier = "Foo()",
                    SelectionModuleName = "Class1",
                    NewName = "Goo"
                };
                AddTestComponent(tdo, tdo.SelectionModuleName, inputCode1, ComponentType.ClassModule);
  
                SetupAndRunRenameRefactorTest(tdo, param);

                var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
                Assert.AreEqual(expectedCode1, rewriter1.GetText());
                tdo.MsgBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()), Times.Never);
            }
        }

        [TestMethod]
        public void RenameRefactoring_RenameControlNameFromEvent_AcceptPrompt()
        {
            const string inputCode1 =
@"
Private Sub cmdBtn1_Click()

End Sub
";

            const string expectedCode1 =
@"
Private Sub bigButton_ClickAgain_Click()

End Sub
";

            var tdo = new RenameTestsDataObject
            {
                SelectionTarget = "cmdBtn1_Click",
                SelectionLineIdentifier = "Private Sub cmdBtn1_Click()",
                SelectionModuleName = "UserForm1",
                NewName = "bigButton_ClickAgain"
            };

            CreateMockVBEForControlsTest(tdo, inputCode1, "cmdBtn1");

            SetupAndRunRenameRefactorTest(tdo, RefactorParams.QualifiedSelection);
            
            var rewriter1 = tdo.ParserState.GetRewriter(RetrieveComponent(tdo, tdo.SelectionModuleName).CodeModule.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, msgbox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter1 = state.GetRewriter(module1.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = state.GetRewriter(module2.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(model.Selection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()), Times.Once);

            var rewriter1 = state.GetRewriter(module1.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = state.GetRewriter(module2.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());
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
            string expectedCode1 = inputCode1;
            string expectedCode2 = inputCode2;

            var selection = new Selection(3, 23, 3, 27);

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
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.No);

            var vbeWrapper = vbe.Object;
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection);

            //SetupFactory
            var factory = SetupFactory(model);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(model.Selection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()), Times.Once);

            var rewriter1 = state.GetRewriter(module1.Parent);
            Assert.AreEqual(expectedCode1, rewriter1.GetText());

            var rewriter2 = state.GetRewriter(module2.Parent);
            Assert.AreEqual(expectedCode2, rewriter2.GetText());
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
            var factory = new RenamePresenterFactory(vbeWrapper, null, state);

            var refactoring = new RenameRefactoring(vbeWrapper, factory, null, state);
            refactoring.Refactor();

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(inputCode, rewriter.GetText());
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
            var factory = new RenamePresenterFactory(vbeWrapper, null, state);

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
            var factory = new RenamePresenterFactory(vbeWrapper, null, state);

            var presenter = factory.Create();
            Assert.AreEqual(null, presenter.Show());
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
            var model = new RenameModel(vbeWrapper, state, default(QualifiedSelection)) { NewName = newName };
            model.Target = model.Declarations.First(i => i.DeclarationType == DeclarationType.Project && i.IsUserDefined);

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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.No);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.Yes);

            var refactoring = new RenameRefactoring(vbeWrapper, factory.Object, messageBox.Object, state);
            refactoring.Refactor(qualifiedSelection);

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
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
            var model = new RenameModel(vbeWrapper, state, qualifiedSelection) { NewName = newName };
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
            presenter.Setup(p => p.Model).Returns(model);
            presenter.Setup(p => p.Show()).Returns(model);
            presenter.Setup(p => p.Show(It.IsAny<Declaration>())).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRenamePresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion

        private void SetupAndRunRenameRefactorTest(RenameTestsDataObject tdo, RefactorParams refactorParam)
        {
            tdo.MsgBox = new Mock<IMessageBox>();
            tdo.MsgBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), MessageBoxButtons.YesNo, It.IsAny<MessageBoxIcon>()))
                  .Returns(DialogResult.Yes);

            if(tdo.VBE == null)
            {
                tdo.VBE = BuildProject(tdo.ProjectName, tdo.Components);
            }
            tdo.ParserState = MockParser.CreateAndParse(tdo.VBE);

            CreateQualifiedSelectionForTestCase(tdo);
            tdo.RenameModel = new RenameModel(tdo.VBE, tdo.ParserState, tdo.QualifiedSelection) { NewName = tdo.NewName };

            //SetupFactory
            var factory = SetupFactory(tdo.RenameModel);

            var refactoring = new RenameRefactoring(tdo.VBE, factory.Object, tdo.MsgBox.Object, tdo.ParserState);
            if(refactorParam == RefactorParams.Declaration)
            {
                refactoring.Refactor(tdo.RenameModel.Target);
            }
            else if(refactorParam == RefactorParams.QualifiedSelection)
            {
                refactoring.Refactor(tdo.QualifiedSelection);
            }
            else
            {
                refactoring.Refactor();
            }
        }

        private void CreateQualifiedSelectionForTestCase(RenameTestsDataObject tdo)
        {
            var component = RetrieveComponent(tdo, tdo.SelectionModuleName);
            var moduleContent = component.CodeModule.GetLines(1, component.CodeModule.CountOfLines);

            var splitToken = new[] { "\r\n" };

            var lines = moduleContent.Split(splitToken, StringSplitOptions.None);
            int lineOfInterestNumber = 0;
            string lineOfInterestContent = string.Empty;
            for (int idx = 0; idx < lines.Length && lineOfInterestNumber < 1; idx++)
            {
                if (lines[idx].Contains(tdo.SelectionLineIdentifier))
                {
                    lineOfInterestNumber = idx + 1;
                    lineOfInterestContent = lines[idx];
                }
            }
            Assert.IsTrue(lineOfInterestNumber > 0, "Unable to find target '" + tdo.SelectionTarget + "' in " + tdo.SelectionModuleName + " content.");
            var column = lineOfInterestContent.IndexOf(tdo.SelectionLineIdentifier, StringComparison.Ordinal);
            column = column + tdo.SelectionLineIdentifier.IndexOf(tdo.SelectionTarget, StringComparison.Ordinal) + 1;

            var moduleParent = component.CodeModule.Parent;
            tdo.QualifiedSelection = new QualifiedSelection(new QualifiedModuleName(moduleParent), new Selection(lineOfInterestNumber, column, lineOfInterestNumber, column));
        }

        private void AddTestComponent(RenameTestsDataObject tdo, string moduleIdentifier, string moduleContent, ComponentType componentType)
        {
            if (null == tdo.Components)
            {
                tdo.Components = new List<TestComponentSpecification>();
            }
            tdo.Components.Add(new TestComponentSpecification(moduleIdentifier, moduleContent, componentType));
        }

        private IVBE BuildProject(string projectName, List<TestComponentSpecification> testComponents)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            testComponents.ForEach(c => enclosingProjectBuilder.AddComponent(c.Name, c.ModuleType, c.Content));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            return builder.Build().Object;
        }

        private IVBComponent RetrieveComponent(RenameTestsDataObject tdo, string componentName)
        {
            var vbProject = tdo.VBE.VBProjects.Single(item => item.Name == tdo.ProjectName);
            return vbProject.VBComponents.SingleOrDefault(item => item.Name == componentName);
        }

        internal class TestComponentSpecification
        {
            public TestComponentSpecification(string componentName, string componentContent, ComponentType componentType)
            {
                Name = componentName;
                Content = componentContent;
                ModuleType = componentType;
            }

            public string Name { get; }
            public string Content { get; }
            public ComponentType ModuleType { get; }
        }


        enum RefactorParams
        {
            None,
            QualifiedSelection,
            Declaration
        };

        internal class RenameTestsDataObject
        {
            public RenameTestsDataObject()
            {
                ProjectName = "TestProject";
            }
            public IVBE VBE { get; set; }
            public RubberduckParserState ParserState { get; set; }
            public List<TestComponentSpecification> Components { get; set; }
            public string ProjectName { get; set; }
            public string NewName { get; set; }
            public string SelectionModuleName { get; set; }
            public string SelectionTarget { get; set; }
            public string SelectionLineIdentifier { get; set; }
            public QualifiedSelection QualifiedSelection { get; set; }
            public RenameModel RenameModel { get; set; }
            public Mock<IMessageBox> MsgBox { get; set; }
        }
    }
}
