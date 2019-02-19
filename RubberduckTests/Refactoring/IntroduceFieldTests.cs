using System;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class IntroduceFieldTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var target = state.AllUserDeclarations.SingleOrDefault(e => e.IdentifierName == "bat");

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(target);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(qualifiedSelection);

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {

                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var messageBox = new Mock<IMessageBox>();

                var refactoring = TestRefactoring(rewritingManager, state, messageBox.Object);
                refactoring.Refactor(qualifiedSelection);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                const string expectedCode = inputCode;
                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var messageBox = new Mock<IMessageBox>();

                var refactoring = TestRefactoring(rewritingManager, state, messageBox.Object);
                refactoring.Refactor(qualifiedSelection);

                messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                const string expectedCode = inputCode;
                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
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
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(component), selection);

                var refactoring = TestRefactoring(rewritingManager, state);
                refactoring.Refactor(state.AllUserDeclarations.FindVariable(qualifiedSelection));

                var actualCode = component.CodeModule.Content();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void IntroduceFieldRefactoring_PassInTarget_Nonvariable()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
Dim bar As Boolean
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var messageBox = new Mock<IMessageBox>();

                var refactoring = TestRefactoring(rewritingManager, state, messageBox.Object);

                try
                {
                    refactoring.Refactor(state.AllUserDeclarations.First(d => d.DeclarationType != DeclarationType.Variable));
                }
                catch (ArgumentException e)
                {
                    messageBox.Verify(m => m.NotifyWarn(It.IsAny<string>(), It.IsAny<string>()), Times.Once);

                    Assert.AreEqual("target", e.ParamName);
                    var actual = component.CodeModule.Content();
                    Assert.AreEqual(inputCode, actual);
                    return;
                }

                Assert.Fail();
            }
        }


        private static IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, IMessageBox msgBox = null)
        {
            var selectionService = MockedSelectionService();
            if (msgBox == null)
            {
                msgBox = new Mock<IMessageBox>().Object;
            }
            return new IntroduceFieldRefactoring(state, msgBox, rewritingManager, selectionService);
        }

        private static ISelectionService MockedSelectionService()
        {
            QualifiedSelection? activeSelection = null;
            var selectionServiceMock = new Mock<ISelectionService>();
            selectionServiceMock.Setup(m => m.ActiveSelection()).Returns(() => activeSelection);
            selectionServiceMock.Setup(m => m.TrySetActiveSelection(It.IsAny<QualifiedSelection>()))
                .Returns(() => true).Callback((QualifiedSelection selection) => activeSelection = selection);
            return selectionServiceMock.Object;
        }
    }
}
