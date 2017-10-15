using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class IntroduceParameterTests
    {
        [TestMethod]
        public void IntroduceParameterRefactoring_NoParamsInList_Sub()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal bar As Boolean)
    
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_NoParamsInList_Function()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Dim bar As Boolean
    Foo = True
End Function";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private Function Foo(ByVal bar As Boolean) As Boolean
    
    Foo = True
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_OneParamInList()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer)
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, ByVal bar As Boolean)
    
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_OneParamInList_MultipleLines()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer)
    Dim _
    bar _
    As _
    Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, ByVal bar As Boolean)
    
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_MultipleParamsOnMultipleLines()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean
End Sub";
            var selection = new Selection(3, 8, 3, 20);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date, ByVal bar As Boolean)
    
End Sub";   // note: the VBE removes extra spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_MultipleVariablesInStatement_MoveFirst()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date, _
        bap As Integer
End Sub";
            var selection = new Selection(3, 10, 3, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date, ByVal bar As Boolean)
    Dim _
        bat As Date, _
        bap As Integer
End Sub";   // note: the VBE removes extra spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_MultipleVariablesInStatement_MoveSecond()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date, _
        bap As Integer
End Sub";
            var selection = new Selection(4, 10, 4, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date, ByVal bat As Date)
    Dim bar As Boolean, _
         _
        bap As Integer
End Sub";   // note: the VBE removes extra spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_MultipleVariablesInStatement_MoveLast()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, _
        bat As Date, _
        bap As Integer
End Sub";
            var selection = new Selection(5, 10, 5, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date, ByVal bap As Integer)
    Dim bar As Boolean, _
        bat As Date
        
End Sub";   // note: the VBE removes extra spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_MultipleVariablesInStatement_OnOneLine_MoveFirst()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date)
    Dim bar As Boolean, bat As Date, bap As Integer
End Sub";
            var selection = new Selection(3, 10, 3, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal buz As Integer, _
                  ByRef baz As Date, ByVal bar As Boolean)
    Dim bat As Date, bap As Integer
End Sub";   // note: the VBE removes extra spaces

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_DisplaysInvalidSelectionAndDoesNothingForField()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean

Private Sub Foo()
End Sub";
            var selection = new Selection(1, 14, 1, 14);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_DisplaysInvalidSelectionAndDoesNothingForInvalidSelection()
        {
            //Input
            const string inputCode =
@"Private fizz As Boolean

Private Sub Foo()
End Sub";
            var selection = new Selection(3, 16, 3, 16);

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                        It.IsAny<MessageBoxIcon>())).Returns(DialogResult.OK);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_Properties_GetAndLet()
        {
            //Input
            const string inputCode =
@"Property Get Foo(ByVal fizz As Boolean) As Boolean
    Dim bar As Integer
    Foo = fizz
End Property

Property Let Foo(ByVal fizz As Boolean, ByVal buzz As Boolean)
End Property";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Property Get Foo(ByVal fizz As Boolean, ByVal bar As Integer) As Boolean
    
    Foo = fizz
End Property

Property Let Foo(ByVal fizz As Boolean, ByVal bar As Integer, ByVal buzz As Boolean)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_Properties_GetAndSet()
        {
            //Input
            const string inputCode =
@"Property Get Foo(ByVal fizz As Boolean) As Variant
    Dim bar As Integer
    Foo = fizz
End Property

Property Set Foo(ByVal fizz As Boolean, ByVal buzz As Variant)
End Property";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Property Get Foo(ByVal fizz As Boolean, ByVal bar As Integer) As Variant
    
    Foo = fizz
End Property

Property Set Foo(ByVal fizz As Boolean, ByVal bar As Integer, ByVal buzz As Variant)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_ImplementsInterface()
        {
            //Input
            const string inputCode1 =
            @"Sub fizz(ByVal boo As Boolean)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean)
    Dim fizz As Date
End Sub";
            var selection = new Selection(4, 10, 4, 14);

            //Expectation
            const string expectedCode1 =
@"Sub fizz(ByVal boo As Boolean, ByVal fizz As Date)
End Sub";

            const string expectedCode2 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean, ByVal fizz As Date)
    
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[1];
            vbe.Setup(v => v.ActiveCodePane).Returns(component.CodeModule.CodePane);

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_ImplementsInterface_MultipleInterfaceImplementations()
        {
            //Input
            const string inputCode1 =
@"Sub fizz(ByVal boo As Boolean)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean)
    Dim fizz As Date
End Sub";

            const string inputCode3 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean)
End Sub";
            var selection = new Selection(4, 10, 4, 14);

            //Expectation
            const string expectedCode1 =
@"Sub fizz(ByVal boo As Boolean, ByVal fizz As Date)
End Sub";

            const string expectedCode2 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean, ByVal fizz As Date)
    
End Sub";

            const string expectedCode3 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean, ByVal fizz As Date)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[1];
            vbe.Setup(v => v.ActiveCodePane).Returns(component.CodeModule.CodePane);

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;
            var module3 = project.Object.VBComponents[2].CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(expectedCode1, module1.Content());
            Assert.AreEqual(expectedCode2, module2.Content());
            Assert.AreEqual(expectedCode3, module3.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_ImplementsInterface_Reject()
        {
            //Input
            const string inputCode1 =
            @"Sub fizz(ByVal boo As Boolean)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_fizz(ByVal boo As Boolean)
    Dim fizz As Date
End Sub";
            var selection = new Selection(4, 10, 4, 14);

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component = project.Object.VBComponents[1];
            vbe.Setup(v => v.ActiveCodePane).Returns(component.CodeModule.CodePane);

            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var module1 = project.Object.VBComponents[0].CodeModule;
            var module2 = project.Object.VBComponents[1].CodeModule;

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.No);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            Assert.AreEqual(inputCode1, module1.Content());
            Assert.AreEqual(inputCode2, module2.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_PassInTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal bar As Boolean)
    
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null);
            refactoring.Refactor(state.AllUserDeclarations.FindVariable(qualifiedSelection));

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceParameterRefactoring_PassInTarget_Nonvariable()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim bar As Boolean
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                      .Returns(DialogResult.OK);

            var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object);

            try
            {
                refactoring.Refactor(state.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));
            }
            catch (ArgumentException e)
            {
                messageBox.Verify(m =>
                    m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                    It.IsAny<MessageBoxIcon>()), Times.Once);

                Assert.AreEqual("target", e.ParamName);
                Assert.AreEqual(inputCode, component.CodeModule.Content());
                return;
            }

            Assert.Fail();
        }
    }
}
