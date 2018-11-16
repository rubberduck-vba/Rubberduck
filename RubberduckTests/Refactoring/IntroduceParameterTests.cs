using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.IntroduceParameter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class IntroduceParameterTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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
Dim bat As Date, _
bap As Integer
End Sub";   // note: the VBE removes extra spaces

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component, selection);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date, ByVal bat As Date)
Dim bar As Boolean, _
bap As Integer
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bat");

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, new Mock<IMessageBox>().Object, rewritingManager);
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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
            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date, ByVal bap As Integer)
Dim bar As Boolean, _
bat As Date
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bap");

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
        public void IntroduceParameterRefactoring_MultipleVariablesInStatement_OnOneLine_MoveFirst()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bar As Boolean, bat As Date, bap As Integer
End Sub";

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date, ByVal bar As Boolean)
Dim bat As Date, bap As Integer
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bar");

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
        public void IntroduceParameterRefactoring_DisplaysInvalidSelectionAndDoesNothingForField()
        {
            //Input
            const string inputCode =
                @"Private fizz As Boolean

Private Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var messageBox = new Mock<IMessageBox>();

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "fizz");
                refactoring.Refactor(target);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                const string expectedCode = inputCode;
                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
        public void IntroduceParameterRefactoring_DisplaysInvalidSelectionAndDoesNothingForInvalidSelection()
        {
            //Input
            const string inputCode =
                @"Private fizz As Boolean

Private Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var messageBox = new Mock<IMessageBox>();

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "fizz");
                refactoring.Refactor(target);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                const string expectedCode = inputCode;
                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            //Expectation
            const string expectedCode =
                @"Property Get Foo(ByVal fizz As Boolean, ByVal bar As Integer) As Boolean
Foo = fizz
End Property

Property Let Foo(ByVal fizz As Boolean, ByVal bar As Integer, ByVal buzz As Boolean)
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bar");
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            //Expectation
            const string expectedCode =
                @"Property Get Foo(ByVal fizz As Boolean, ByVal bar As Integer) As Variant
Foo = fizz
End Property

Property Set Foo(ByVal fizz As Boolean, ByVal bar As Integer, ByVal buzz As Variant)
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bar");
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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
            var component1 = project.Object.VBComponents[0];
            var component2 = project.Object.VBComponents[1];
            vbe.Setup(v => v.ActiveCodePane).Returns(component1.CodeModule.CodePane);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m => m.Question(It.IsAny<string>(), It.IsAny<string>())).Returns(true);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "fizz" && e.DeclarationType == DeclarationType.Variable);
                refactoring.Refactor(target);

                var actualCode1 = component1.CodeModule.Content();
                Assert.AreEqual(expectedCode1, actualCode1);

                var actualCode2 = component2.CodeModule.Content();
                Assert.AreEqual(expectedCode2, actualCode2);

                messageBox.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Once());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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
            var component1 = project.Object.VBComponents[0];
            var component2 = project.Object.VBComponents[1];
            var component3 = project.Object.VBComponents[2];
            vbe.Setup(v => v.ActiveCodePane).Returns(component2.CodeModule.CodePane);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m => m.Question(It.IsAny<string>(), It.IsAny<string>())).Returns(true);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "fizz" && e.DeclarationType == DeclarationType.Variable);
                refactoring.Refactor(target);

                var actualCode1 = component1.CodeModule.Content();
                Assert.AreEqual(expectedCode1, actualCode1);

                var actualCode2 = component2.CodeModule.Content();
                Assert.AreEqual(expectedCode2, actualCode2);

                var actualCode3 = component3.CodeModule.Content();
                Assert.AreEqual(expectedCode3, actualCode3);

                messageBox.Verify(m => m.Question(It.IsAny<string>(), It.IsAny<string>()), Times.Once());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var component1 = project.Object.VBComponents[0];
            var component2 = project.Object.VBComponents[1];
            vbe.Setup(v => v.ActiveCodePane).Returns(component2.CodeModule.CodePane);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var messageBox = new Mock<IMessageBox>();
                messageBox.Setup(m => m.ConfirmYesNo(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<bool>())).Returns(false);

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "fizz" && e.DeclarationType == DeclarationType.Variable);
                refactoring.Refactor(target);

                const string expectedCode1 = inputCode1;
                var actualCode1 = component1.CodeModule.Content();
                Assert.AreEqual(expectedCode1, actualCode1);

                const string expectedCode2 = inputCode2;
                var actualCode2 = component2.CodeModule.Content();
                Assert.AreEqual(expectedCode2, actualCode2);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
        public void IntroduceParameterRefactoring_PassInTarget()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
Dim bar As Boolean
End Sub";

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal bar As Boolean)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, null, rewritingManager);

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bar" && e.DeclarationType == DeclarationType.Variable);
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Parameter")]
        public void IntroduceParameterRefactoring_PassInTarget_Nonvariable()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
Dim bar As Boolean
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out IVBComponent component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var messageBox = new Mock<IMessageBox>();

                var refactoring = new IntroduceParameterRefactoring(vbe.Object, state, messageBox.Object, rewritingManager);
                refactoring.Refactor(state.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                const string expectedCode = inputCode;
                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }
    }
}
