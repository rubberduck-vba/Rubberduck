using System;
using NUnit.Framework;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.Utility;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;
using Rubberduck.Refactorings.Exceptions.ExtractMethod;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestFixture]
    public class ExtractMethodTests : InteractiveRefactoringTestBase<IExtractMethodPresenter, ExtractMethodModel>
    {
        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void InboundOnlyWithoutPreassignmentCopiesDeclaration()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
End Sub";
            var selection = new Selection(4, 5, 4, 11);
            var expectedNewMethodCode = @"
Private Sub NewMethod()
    Dim a As Integer
    a = 10
End Sub";
            var expectedReplacementCode = @"
    NewMethod";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;
            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void InboundOnlyWithPreassignmentPassesByVal()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
    Debug.Print a
End Sub";
            var selection = new Selection(5, 5, 5, 18);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByVal a As Integer)
    Debug.Print a
End Sub";
            var expectedReplacementCode = @"
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void InboundAndOutboundPassByRef()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
    a = a + 10
    Debug.Print a
End Sub";
            var selection = new Selection(5, 5, 5, 15);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    a = a + 10
End Sub";
            var expectedReplacementCode = @"
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void ParameterlessWorks()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
    a = a + 10
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 6, 18);
            var expectedNewMethodCode = @"
Private Sub NewMethod()
    Dim a As Integer
    a = 10
    a = a + 10
    Debug.Print a
End Sub";
            var expectedReplacementCode = @"
    NewMethod";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void OutboundOnlyMovesDeclaration()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
    a = a + 10
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 4, 11);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    a = 10
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void InboundOnlyWithoutPreassignmentForObject()
        {
            var inputCode = @"
Sub Test
    Dim a As Object
    Set a = New Collection
End Sub";
            var selection = new Selection(4, 5, 4, 27);
            var expectedNewMethodCode = @"
Private Sub NewMethod()
    Dim a As Object
    Set a = New Collection
End Sub";
            var expectedReplacementCode = @"
    NewMethod";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;
            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void InboundOnlyWithPreassignmentForObject()
        {
            var inputCode = @"
Sub Test
    Dim a As Object
    Set a = New Collection
    a.Add 1
    Debug.Print a(1)
End Sub";
            var selection = new Selection(6, 5, 6, 21);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByVal a As Object)
    Debug.Print a(1)
End Sub";
            var expectedReplacementCode = @"
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void InboundAndOutboundForObject()
        {
            var inputCode = @"
Sub Test
    Dim a As New Collection
    a.Add 1
    a(1) = a(1) + 10
    Debug.Print a(1)
End Sub";
            var selection = new Selection(5, 5, 5, 21);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Collection)
    a(1) = a(1) + 10
End Sub";
            var expectedReplacementCode = @"
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void ParameterlessForObject()
        {
            var inputCode = @"
Sub Test
    Dim a As New Collection
    a.Add 1
    a(1) = a(1) + 10
    Debug.Print a(1)
End Sub";
            var selection = new Selection(3, 1, 6, 21);
            var expectedNewMethodCode = @"
Private Sub NewMethod()
    Dim a As New Collection
    a.Add 1
    a(1) = a(1) + 10
    Debug.Print a(1)
End Sub";
            var expectedReplacementCode = @"
    NewMethod";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void OutboundOnlyForObject()
        {
            var inputCode = @"
Sub Test
    Dim a As Object
    Set a = New Collection
    a.Add 1
    a(1) = a(1) + 10
    Debug.Print a(1)
End Sub";
            var selection = new Selection(3, 1, 5, 12);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Object)
    Set a = New Collection
    a.Add 1
End Sub";
            var expectedReplacementCode = @"
    Dim a As Object
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void FunctionWorksForValueType()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
    Debug.Print a
End Sub";
            var selection = new Selection(4, 5, 4, 11);
            var expectedNewMethodCode = @"
Private Function NewMethod() As Integer
    Dim a As Integer
    a = 10

    NewMethod = a
End Function";
            var expectedReplacementCode = @"
    a = NewMethod()";
            var model = InitialModel(inputCode, selection, true);
            model.ReturnParameter = model.Parameters[0];
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void FunctionWorksForObjectType()
        {
            var inputCode = @"
Sub Test
    Dim a As New Collection
    a.Add 10
    Debug.Print a.Count
End Sub";
            var selection = new Selection(3, 1, 4, 13);
            var expectedNewMethodCode = @"
Private Function NewMethod() As Collection
    Dim a As New Collection
    a.Add 10

    Set NewMethod = a
End Function";
            var expectedReplacementCode = @"
    Dim a As New Collection
    Set a = NewMethod()";
            var model = InitialModel(inputCode, selection, true);
            model.ReturnParameter = model.Parameters[0];
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test(Description = "Inbound only reference normally would be ByVal but not allowed for array variables")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void ArrayCannotBeByVal()
        {
            var inputCode = @"
Sub Test
    Dim a(0 To 0) As Integer
    a(0) = 10
    Debug.Print a(0)
End Sub";
            var selection = new Selection(5, 5, 5, 21);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a() As Integer)
    Debug.Print a(0)
End Sub";
            var expectedReplacementCode = @"
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test(Description = "Inbound only reference normally would be ByVal but not allowed for array variables")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void AvoidNewMethodNameClash()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer
    a = 10
    Debug.Print a
End Sub

Sub AnotherSub
End Sub

Sub AnotherSub1
End Sub";
            var selection = new Selection(4, 1, 4, 11);
            var expectedNewMethodCode = @"
Private Sub AnotherSub2(ByRef a As Integer)
    a = 10
End Sub";
            var expectedReplacementCode = @"
    AnotherSub2 a";
            var model = InitialModel(inputCode, selection, true);
            model.NewMethodName = "AnotherSub";
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void ThrowsWhenNeedToMoveMultivariableDeclaration()
        {
            var inputCode = @"
Sub Test
    Dim b As Integer, a As Integer
    a = 10
    a = a + 10
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 4, 11);
            var model = InitialModel(inputCode, selection, true);
            Assert.Throws(typeof(UnableToMoveVariableDeclarationException), () => { var _ = model.NewMethodCode; } );
        }


        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void ThrowsWhenModifyingReturnValueOfFunction()
        {
            var inputCode = @"
Function Test(ByVal i As Integer) As Integer
    Dim a As Integer
    if i > 0 Then
        i = i - 1
        a = Test(i) 'Recursive type call
    End If
    Test = 10 'Setting return value complicates things
End Function";
            var selection = new Selection(4, 1, 8, 55);
            Assert.Throws(typeof(InvalidTargetSelectionException), () => {
                var model = InitialModel(inputCode, selection, true);
            });
        }

        [Test(Description = "Handle indenting and splitting of code when two statements in one colon-separated line")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void FindCorrectIndentingForReplacementCode()
        {
            var inputCode = @"
Sub Test
    Dim a As Integer: Dim b As Integer
    a = 10
    b = 12
    Debug.Print a + b
End Sub";
            var selection = new Selection(3, 23, 6, 22);
            var expectedNewMethodCode = @"
Private Sub NewMethod()
    Dim a As Integer
    Dim b As Integer
    a = 10
    b = 12
    Debug.Print a + b
End Sub";
            var expectedReplacementCode = @"
    NewMethod";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test(Description = "Two references to a variable in the same line, only one extracted")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void SplitOutLogicalStatementFromMultistatementLine()
        {
            var inputCode = @"
Sub splitter()
    Dim a As Integer
    Dim b As Integer
    b = 10: b = b + 10
    a = b
    Debug.Print a
End Sub";
            var selection = new Selection(5, 13, 6, 10);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByVal b As Integer, ByRef a As Integer)
    b = b + 10
    a = b
End Sub";
            var expectedReplacementCode = @"
    NewMethod b, a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void OutboundOnlyForDimAsNewObject()
        {
            var inputCode = @"
Sub Test
    Dim a As New Collection
    a.Add 1
    Debug.Print a(1)
End Sub";
            var selection = new Selection(3, 1, 4, 12);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Collection)
    a.Add 1
End Sub";
            var expectedReplacementCode = @"
    Dim a As New Collection
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test(Description = "Selected comments at start and end of selection get copied")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void CommentsGetCopiedToNewMethod()
        {
            var inputCode = @"
Sub CopiesComment()
    'Start comment
    Dim a As Integer
    a = 1
    'End comment
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 6, 17);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    'Start comment
    a = 1
    'End comment
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test(Description = "Preserve statement preceeding the declaration to move")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void DeclarationToMoveAfterStatementOnSameLine()
        {
            var inputCode = @"
Sub Test()
    Debug.Print 0: Dim a As Integer
    a = 1
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 4, 10);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    Debug.Print 0
    a = 1
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test(Description = "Preserve statement following the declaration to move")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void DeclarationToMoveBeforeStatementOnSameLine()
        {
            var inputCode = @"
Sub Test()
    Dim a As Integer: Debug.Print 0
    a = 1
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 4, 10);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    Debug.Print 0
    a = 1
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test(Description = "Preserve statements surrounding the declaration to move")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void DeclarationToMoveExtractedFromMiddleOfLine()
        {
            var inputCode = @"
Sub Test()
    Debug.Print 0: Dim a As Integer: Debug.Print 1
    a = 1
    Debug.Print a
End Sub";
            var selection = new Selection(3, 1, 4, 10);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    Debug.Print 0: Debug.Print 1
    a = 1
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        [Test(Description = "Removing a declaration from a partially selected line with other code as well")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void DeclarationToMoveExtractedFromPartiallySelectedMultiStatementLine()
        {
            var inputCode = @"
Sub Test()
    Debug.Print 0: Dim a As Integer: Debug.Print 1
    a = 1
    Debug.Print a
End Sub";
            var selection = new Selection(3, 20, 4, 10);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    Debug.Print 1
    a = 1
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }


        [Test(Description = "Cut parts from a single multi-statement line and move declaration")]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        public void DeclarationToMoveExtractedFromPartiallySelectedMultiStatementSingleLine()
        {
            var inputCode = @"
Sub Test()
    Debug.Print 0: Dim a As Integer: a = 0: Debug.Print 1
    a = 1
    Debug.Print a
End Sub";
            var selection = new Selection(3, 20, 3, 43);
            var expectedNewMethodCode = @"
Private Sub NewMethod(ByRef a As Integer)
    a = 0
End Sub";
            var expectedReplacementCode = @"
    Dim a As Integer
    NewMethod a";
            var model = InitialModel(inputCode, selection, true);
            var newMethodCode = model.NewMethodCode;

            Assert.AreEqual(expectedNewMethodCode, newMethodCode);
            var replacementCode = model.ReplacementCode;
            Assert.AreEqual(expectedReplacementCode, replacementCode);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IExtractMethodPresenter, ExtractMethodModel> userInteraction,
            ISelectionService selectionService)
        {
            var refactoringAction = new ExtractMethodRefactoringAction(rewritingManager);
            var indenter = new Indenter(null, () => IndenterSettingsTests.GetMockIndenterSettings());
            var selectedDeclarationService = new SelectedDeclarationProvider(selectionService, state);
            return new ExtractMethodRefactoring(refactoringAction, state, userInteraction, selectionService, selectedDeclarationService, state?.ProjectsProvider, indenter);
        }
    }
}