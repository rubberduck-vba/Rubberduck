using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class IntroduceFieldTests
    {
        [TestMethod]
        public void IntroduceFieldRefactoring_NoFieldsInClass_Sub()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_NoFieldsInList_Function()
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
@"Private bar As Boolean
Private Function Foo() As Boolean
Foo = True
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_OneFieldInList()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Private Sub Foo(ByVal buz As Integer)
Dim bar As Boolean
End Sub";
            var selection = new Selection(3, 10, 3, 13);

            //Expectation
            const string expectedCode =
@"Public fizz As Integer
Private bar As Boolean
Private Sub Foo(ByVal buz As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_OneFieldInList_MultipleLines()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Private Sub Foo(ByVal buz As Integer)
Dim _
bar _
As _
Boolean
End Sub";
            var selection = new Selection(3, 10, 3, 13);

            //Expectation
            const string expectedCode =
@"Public fizz As Integer
Private bar As Boolean
Private Sub Foo(ByVal buz As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_MultipleFieldsOnMultipleLines()
        {
            //Input
            const string inputCode =
@"Public fizz As Integer
Public buzz As Integer
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bar As Boolean
End Sub";
            var selection = new Selection(5, 8, 5, 20);

            //Expectation
            const string expectedCode =
@"Public fizz As Integer
Public buzz As Integer
Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_MultipleVariablesInStatement_MoveFirst()
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
@"Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bat As Date, _
bap As Integer
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_MultipleVariablesInStatement_MoveSecond()
        {
            //Input
            const string inputCode = @"
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bar As Boolean, _
bat As Date, _
bap As Integer
End Sub";
            //Expectation
            const string expectedCode = @"
Private bat As Date
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bar As Boolean, _
bap As Integer
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bat");

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, new Mock<IMessageBox>().Object);
            refactoring.Refactor(target);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_MultipleVariablesInStatement_MoveLast()
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
@"Private bap As Integer
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bar As Boolean, _
bat As Date
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_MultipleVariablesInStatement_OnOneLine_MoveFirst()
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
@"Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bat As Date, bap As Integer
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, null);
            refactoring.Refactor(qualifiedSelection);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_DisplaysInvalidSelectionAndDoesNothingForField()
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

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);
            Assert.AreEqual(inputCode, component.CodeModule.Content());
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_DisplaysInvalidSelectionAndDoesNothingForInvalidSelection()
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

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, messageBox.Object);
            refactoring.Refactor(qualifiedSelection);

            messageBox.Verify(m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(),
                It.IsAny<MessageBoxIcon>()), Times.Once);

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(inputCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_PassInTarget()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
@"Private bar As Boolean
Private Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component, selection);
            var state = MockParser.CreateAndParse(vbe.Object);

            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

            var refactoring = new IntroduceFieldRefactoring((vbe.Object), state, null);
            refactoring.Refactor(state.AllUserDeclarations.FindVariable(qualifiedSelection));

            var actual = state.GetRewriter(component).GetText();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void IntroduceFieldRefactoring_PassInTarget_Nonvariable()
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

            var refactoring = new IntroduceFieldRefactoring(vbe.Object, state, messageBox.Object);

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
                var actual = state.GetRewriter(component).GetText();
                Assert.AreEqual(inputCode, actual);
                return;
            }

            Assert.Fail();
        }
    }
}
