using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.IntroduceField;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.IntroduceField;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class IntroduceFieldTests : RefactoringTestBase
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

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void IntroduceFieldRefactoring_NoFieldsInClass_MultipleSub()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
Dim bar As Boolean
End Sub

Private Sub Baz()
Dim bar As Boolean
End Sub";
            var selection = new Selection(2, 10, 2, 13);

            //Expectation
            const string expectedCode =
                @"Private bar As Boolean
Private Sub Foo()
End Sub

Private Sub Baz()
Dim bar As Boolean
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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
            var selection = new Selection(3, 5, 3, 8);

            //Expectation
            const string expectedCode =
                @"Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bat As Date, _
bap As Integer
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = RefactoredCode(inputCode, "bat", DeclarationType.Variable);
            Assert.AreEqual(expectedCode, actualCode);
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
            var selection = new Selection(5, 1, 5, 4);

            //Expectation
            const string expectedCode =
                @"Private bap As Integer
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bar As Boolean, _
bat As Date
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
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
            var selection = new Selection(3, 5, 3, 7);

            //Expectation
            const string expectedCode =
                @"Private bar As Boolean
Private Sub Foo(ByVal buz As Integer, _
ByRef baz As Date)
Dim bat As Date, bap As Integer
End Sub";

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void IntroduceFieldRefactoring_ThrowsTargetDeclarationIsAlreadyAFieldExceptionAndDoesNothingForField()
        {
            //Input
            const string inputCode =
                @"Private fizz As Boolean

Private Sub Foo()
End Sub";
            var selection = new Selection(1, 14, 1, 14);

            var actualCode = RefactoredCode(inputCode, selection, typeof(TargetIsAlreadyAFieldException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void IntroduceFieldRefactoring_ThrowsNoDeclarationForSelectionAndDoesNothingForInvalidSelection()
        {
            //Input
            const string inputCode =
                @"Private fizz As Boolean

Private Sub Foo()
End Sub";
            var selection = new Selection(3, 16, 3, 16);

            var actualCode = RefactoredCode(inputCode, selection, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(inputCode, actualCode);
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

            var actualCode = RefactoredCode(inputCode, selection);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Introduce Field")]
        public void IntroduceFieldRefactoring_PassInTarget_NonVariable()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
Dim bar As Boolean
End Sub";

            var actualCode = RefactoredCode(inputCode, "Foo", DeclarationType.Procedure, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, ISelectionService selectionService)
        {
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            var baseRefactoring = new IntroduceFieldRefactoringAction(state, rewritingManager);
            return new IntroduceFieldRefactoring(baseRefactoring, selectionService, selectedDeclarationProvider);
        }
    }
}
