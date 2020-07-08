using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ReorderParameters;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.UI.Refactorings.ReorderParameters;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ReorderParametersTests :InteractiveRefactoringTestBase<IReorderParametersPresenter, ReorderParametersModel>
    {
        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_SwapPositions()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_SwapPositions_SignatureContainsParamName()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(a, ba)
End Sub";
            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ba, a)
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_SwapPositions_ReferenceValueContainsOtherReferenceValue()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(a, ba)
End Sub

Sub Goo()
    Foo 1, 121
End Sub";
            var selection = new Selection(1, 16, 1, 16);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ba, a)
End Sub

Sub Goo()
    Foo 121, 1
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_RefactorDeclaration()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_RefactorDeclaration_FailsInvalidTarget()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, "TestModule1", DeclarationType.ProceduralModule, presenterAction, typeof(InvalidDeclarationTypeException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_RefactorDeclaration_FailsNoValidTargetSelected()
        {
            //Input
            const string inputCode =
                @"Private bar As Long 

Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";
            var selection = Selection.Home;

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction, typeof(NoDeclarationForSelectionException));
            Assert.AreEqual(inputCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_WithOptionalParam()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub";

            var presenterAction = ReorderParamIndices(new List<int>{1, 0, 2});
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_WithOptionalParams_RemovedTrailingMissingArguments()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(Optional ByVal arg1 As Integer = 0, Optional ByVal arg2 As String = vbNullString, Optional ByVal arg3 As Long = 0)
End Sub

Public Sub Goo()
    Foo ,, 6
    Foo , 4, 6
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Public Sub Foo(Optional ByVal arg3 As Long = 0, Optional ByVal arg2 As String = vbNullString, Optional ByVal arg1 As Integer = 0)
End Sub

Public Sub Goo()
    Foo 6
    Foo 6, 4
End Sub
";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_SwapPositions_UpdatesCallers()
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
                @"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Bar()
    Foo ""Hello"", 10
End Sub
";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void RemoveParametersRefactoring_ClientReferencesAreUpdated_ParensAroundCall()
        {
            //Input
            const string inputCode =
                @"Private Sub bar()
    Dim x As Integer
    Dim y As Integer
    y = foo(x, 42)
    Debug.Print y, x
End Sub

Private Function foo(ByRef a As Integer, ByVal b As Integer) As Integer
    a = b
    foo = a + b
End Function";
            var selection = new Selection(8, 20, 8, 20);

            //Expectation
            const string expectedCode =
                @"Private Sub bar()
    Dim x As Integer
    Dim y As Integer
    y = foo(42, x)
    Debug.Print y, x
End Sub

Private Function foo(ByVal b As Integer, ByRef a As Integer) As Integer
    a = b
    foo = a + b
End Function";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_DoesNotReorderNamedArguments()
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
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg3 As Double, ByVal arg2 As String)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg3:=6.1, arg1:=3
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int>{0, 2, 1});
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_MakesPartiallyNamedArgumentsAllNamed()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Double, ByVal arg4 As Double)
End Sub

Public Sub Goo()
    Foo 23, ""test42"", arg4:=6.1, arg3:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Public Sub Foo(ByVal arg4 As Double, ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg1:=23, arg2:=""test42"", arg4:=6.1, arg3:=3
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int> { 3, 0, 1, 2 });
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_DoesNotReorderNamedArguments_Function()
        {
            //Input
            const string inputCode =
                @"Public Function Foo(ByVal arg1 As Integer, ByVal arg2 As String) As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Public Function Foo(ByVal arg2 As String, ByVal arg1 As Integer) As Boolean
    Foo = True
End Function";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_DoesNotReorderNamedArguments_WithOptionalParam()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Public Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Double)
End Sub

Public Sub Goo()
    Foo arg2:=""test44"", arg1:=3
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int>{1, 0, 2});
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_MakesPartiallyNamedArgumentsAllNamedAndDropsMissingOptionalArguments()
        {
            //Input
            const string inputCode =
                @"Public Sub Foo(Optional ByVal arg1 As Integer = 2, Optional ByVal arg2 As String = vbNullString, Optional ByVal arg3 As Double = 1, Optional ByVal arg4 As Double = 0)
End Sub

Public Sub Goo()
    Foo 23, , arg4:=6.1, arg3:=3
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Public Sub Foo(Optional ByVal arg4 As Double = 0, Optional ByVal arg1 As Integer = 2, Optional ByVal arg2 As String = vbNullString, Optional ByVal arg3 As Double = 1)
End Sub

Public Sub Goo()
    Foo arg1:=23, arg4:=6.1, arg3:=3
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int> { 3, 0, 1, 2 });
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderGetter()
        {
            //Input
            const string inputCode =
                @"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date) As Boolean
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Get Foo(ByVal arg2 As String, ByVal arg3 As Date, ByVal arg1 As Integer) As Boolean
End Property";

            var presenterAction = ReorderParamIndices(new List<int> { 1, 2, 0 });
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderLetter()
        {
            //Input
            const string inputCode =
                @"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderSetter()
        {
            //Input
            const string inputCode =
                @"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderLastParamFromSetter_NotAllowed()
        {
            //Input
            const string inputCode =
                @"Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            ReorderParametersModel capturedModel = null;
            Func<ReorderParametersModel, ReorderParametersModel> presenterAction = model =>
            {
                capturedModel = model;
                return model;
            };

            RefactoredCode(inputCode, selection, presenterAction);

            Assert.AreEqual(1, capturedModel.Parameters.Count); // doesn't allow removing last param from setter
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderLastParamFromLetter_NotAllowed()
        {
            //Input
            const string inputCode =
                @"Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 23, 1, 27);

            ReorderParametersModel capturedModel = null;
            Func<ReorderParametersModel, ReorderParametersModel> presenterAction = model =>
            {
                capturedModel = model;
                return model;
            };

            RefactoredCode(inputCode, selection, presenterAction);

            Assert.AreEqual(1, capturedModel.Parameters.Count); // doesn't allow removing last param from setter
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_SignatureOnMultipleLines()
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
                @"Private Sub Foo(ByVal arg3 As Date, _
                  ByVal arg2 As String, _
                  ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_CallOnMultipleLines()
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
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg3 As Date, ByVal arg2 As String, ByVal arg1 As Integer)
End Sub

Private Sub Goo(ByVal arg1 as Integer, ByVal arg2 As String, ByVal arg3 As Date)

    Foo arg3, _
        arg2, _
        arg1

End Sub
";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ParamArrayBeforeOtherArgument_CannotConfirmDialog()
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

            ReorderParametersModel capturedModel = null;
            Func<ReorderParametersModel, ReorderParametersModel> presenterAction = model =>
            {
                capturedModel = model;
                return ReverseParameters()(model);
            };
            RefactoredCode(inputCode, selection, presenterAction);

            var declarationFinderProvider = new Mock<IDeclarationFinderProvider>().Object;
            var viewModel = new ReorderParametersViewModel(declarationFinderProvider, capturedModel);

            Assert.IsFalse(viewModel.OkButtonCommand.CanExecute(null));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ClientReferencesAreUpdated_ParamArray()
        {
            //Input
            const string inputCode =
                @"Sub Foo(ByVal arg1 As String, ByVal arg2 As Date, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", arg, test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Sub Foo(ByVal arg2 As Date, ByVal arg1 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo arg, ""test"", test1x, test2x, test3x, test4x, test5x, test6x
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int>{1, 0, 2});
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ClientReferencesAreUpdated_ParamArray_CallOnMultipleLines()
        {
            //Input
            const string inputCode =
                @"Sub Foo(ByVal arg1 As String, ByVal arg2 As Date, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo ""test"", _
        arg, _
        test1x, _
        test2x, _
        test3x, _
        test4x, _
        test5x, _
        test6x
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Sub Foo(ByVal arg2 As Date, ByVal arg1 As String, ParamArray arg3())
End Sub

Public Sub Goo(ByVal arg As Date, _
               ByVal arg1 As Integer, _
               ByVal arg2 As Integer, _
               ByVal arg3 As Integer, _
               ByVal arg4 As Integer, _
               ByVal arg5 As Integer, _
               ByVal arg6 As Integer)
              
    Foo arg, _
        ""test"", _
        test1x, _
        test2x, _
        test3x, _
        test4x, _
        test5x, _
        test6x
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int> { 1, 0, 2 });
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_MoveOptionalParamBeforeNonOptionalParam_CannotConfirmDialog()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, Optional ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub";
            var selection = new Selection(1, 23, 1, 27);

            ReorderParametersModel capturedModel = null;
            Func<ReorderParametersModel, ReorderParametersModel> presenterAction = model =>
            {
                capturedModel = model;
                return ReorderParamIndices(new List<int> {1, 2, 0})(model);
            };
            RefactoredCode(inputCode, selection, presenterAction);

            var declarationFinderProvider = new Mock<IDeclarationFinderProvider>().Object;
            var viewModel = new ReorderParametersViewModel(declarationFinderProvider, capturedModel);

            Assert.IsFalse(viewModel.OkButtonCommand.CanExecute(null));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_ReorderCallsWithoutOptionalParams()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg1, arg2
End Sub
";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Sub Foo(ByVal arg2 As String, ByVal arg1 As Integer, Optional ByVal arg3 As Boolean = True)
End Sub

Private Sub Goo(ByVal arg1 As Integer, ByVal arg2 As String)
    Foo arg2, arg1
End Sub
";

            var presenterAction = ReorderParamIndices(new List<int> { 1, 0, 2 });
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderFirstParamFromGetterAndSetter()
        {
            //Input
            const string inputCode =
                @"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Set Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ReorderFirstParamFromGetterAndLetter()
        {
            //Input
            const string inputCode =
                @"Private Property Get Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Property

Private Property Let Foo(ByVal arg1 As Integer, ByVal arg2 As String, ByVal arg3 As Date)
End Property";
            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedCode =
                @"Private Property Get Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Property

Private Property Let Foo(ByVal arg2 As String, ByVal arg1 As Integer, ByVal arg3 As Date)
End Property";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(inputCode, selection, presenterAction);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParams_PresenterIsNull()
        {
            //Input
            const string inputCode =
                @"Private Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using(state)
            {
                var qualifiedSelection = new QualifiedSelection(component.QualifiedModuleName, Selection.Home);
                var factory = new Mock<IRefactoringPresenterFactory>();
                factory.Setup(f => f.Create<IReorderParametersPresenter, ReorderParametersModel>(It.IsAny<ReorderParametersModel>()))
                    .Returns(() => null); // resolves ambiguous method resolution
                var selectionService = MockedSelectionService(qualifiedSelection);

                var refactoring = TestRefactoring(rewritingManager, state, factory.Object, selectionService);

                Assert.Throws<InvalidRefactoringPresenterException>(() => refactoring.Refactor(qualifiedSelection));

                Assert.AreEqual(inputCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped()
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
                @"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_ParamsHaveDifferentNames()
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
                @"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_ParamsHaveDifferentNames_TwoImplementations()
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
                @"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces
            const string expectedCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["IClass1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class1"]);
            Assert.AreEqual(expectedCode3, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_InterfaceParamsSwapped_AcceptPrompt()
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

            //Expectation
            const string expectedCode1 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
                @"Public Sub DoSomething(ByVal b As String, ByVal a As Integer)
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("IClass1", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["IClass1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Remove Parameters")]
        public void ReorderParametersRefactoring_InterfaceGetterParam_ImplementationLetAndSetParamReordered()
        {
            //Input
            const string interfaceCode =
                @"Public Property Get Foo(ByVal a As Integer, ByVal b As String) As Variant
End Property

Public Property Let Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property

Public Property Set Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property";
            const string implementerCode =
                @"Implements IClass1

Private Property Get IClass1_Foo(ByVal a As Integer, ByVal b As String) As Variant
End Property

Private Property Let IClass1_Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property

Private Property Set IClass1_Foo(ByVal a As Integer, ByVal b As String, RHS As Variant)
End Property";

            var selection = new Selection(1, 23, 1, 27);

            //Expectation
            const string expectedInterfaceCode =
                @"Public Property Get Foo(ByVal b As String, ByVal a As Integer) As Variant
End Property

Public Property Let Foo(ByVal b As String, ByVal a As Integer, RHS As Variant)
End Property

Public Property Set Foo(ByVal b As String, ByVal a As Integer, RHS As Variant)
End Property";
            const string expectedImplementerCode =
                @"Implements IClass1

Private Property Get IClass1_Foo(ByVal b As String, ByVal a As Integer) As Variant
End Property

Private Property Let IClass1_Foo(ByVal b As String, ByVal a As Integer, RHS As Variant)
End Property

Private Property Set IClass1_Foo(ByVal b As String, ByVal a As Integer, RHS As Variant)
End Property";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "IClass1",
                selection,
                presenterAction,
                null,
                false,
                ("IClass1", interfaceCode, ComponentType.ClassModule),
                ("Class1", implementerCode, ComponentType.ClassModule));

            Assert.AreEqual(expectedInterfaceCode, actualCode["IClass1"]);
            Assert.AreEqual(expectedImplementerCode, actualCode["Class1"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_EventParamsSwapped()
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
                @"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_EventParamsSwapped_EventImplementationSelected()
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

Private Sub abc_Foo(ByVal arg2 As String, ByVal arg1 As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode2 =
                @"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_EventParamsSwapped_DifferentParamNames()
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
                @"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_EventParamsSwapped_DifferentParamNames_TwoHandlers()
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
                @"Public Event Foo(ByVal arg2 As String, ByVal arg1 As Integer)";

            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal s As String, ByVal i As Integer)
End Sub";   // note: IDE removes excess spaces

            const string expectedCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal v2 As String, ByVal v1 As Integer)
End Sub";   // note: IDE removes excess spaces

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "Class1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule));

            Assert.AreEqual(expectedCode1, actualCode["Class1"]);
            Assert.AreEqual(expectedCode2, actualCode["Class2"]);
            Assert.AreEqual(expectedCode3, actualCode["Class3"]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("Reorder Parameters")]
        public void ReorderParametersRefactoring_ChangesCorrectReferenceArgumentList()
        {
            //Input
            const string classCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string moduleCode =
                @"
Private Function Bar(ByVal i As Integer, ByVal s As String) As Class1
End Function

Private Sub Baz()
    Bar(42, ""Hello"").Foo 23, ""Hi""
End Sub";

            var selection = new Selection(2, 20, 2, 20);

            //Expectation
            const string expectedClassCode =
                @"Public Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            const string expectedModuleCode =
                @"
Private Function Bar(ByVal s As String, ByVal i As Integer) As Class1
End Function

Private Sub Baz()
    Bar(""Hello"", 42).Foo 23, ""Hi""
End Sub";

            var presenterAction = ReverseParameters();
            var actualCode = RefactoredCode(
                "Module1",
                selection,
                presenterAction,
                null,
                false,
                ("Class1", classCode, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            Assert.AreEqual(expectedClassCode, actualCode["Class1"]);
            Assert.AreEqual(expectedModuleCode, actualCode["Module1"]);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager, 
            RubberduckParserState state,
            RefactoringUserInteraction<IReorderParametersPresenter, ReorderParametersModel> userInteraction, 
            ISelectionService selectionService)
        {
            var selectedDeclarationProvider = new SelectedDeclarationProvider(selectionService, state);
            var baseRefactoring = new ReorderParameterRefactoringAction(state, rewritingManager);
            return new ReorderParametersRefactoring(baseRefactoring, state, userInteraction, selectionService, selectedDeclarationProvider);
        }

        private static Func<ReorderParametersModel, ReorderParametersModel> ReverseParameters()
        {
            return model =>
            {
                model.Parameters.Reverse();
                return model;
            };
        }

        private static Func<ReorderParametersModel, ReorderParametersModel> ReorderParamIndices(IList<int> newParamIndexOrder)
        {
            if (newParamIndexOrder == null)
            {
                return model => model;
            }

            return model =>
            {
                var newParamOrder = newParamIndexOrder
                    .Select(idx => model.Parameters[idx])
                    .ToList();

                model.Parameters = newParamOrder;
                return model;
            };
        }
    }
}
