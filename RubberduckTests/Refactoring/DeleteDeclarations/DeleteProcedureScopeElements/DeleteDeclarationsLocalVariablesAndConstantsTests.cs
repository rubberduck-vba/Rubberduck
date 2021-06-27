using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddInterfaceImplementations;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using RubberduckTests.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class DeleteDeclarationsLocalVariablesAndConstantsTests : DeleteDeclarationsLocalsTestsBase
    {
        [TestCase("Const bizz As Integer = 9")]
        [TestCase("Dim bizz As Integer")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeleteSingleDeclaration(string declaration)
        {
            var inputCode =
$@"
Public Sub Foo()
    {declaration}
End Sub";

            var expected =
$@"
Public Sub Foo()
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Const bizz1 As Integer = 9, bizz2 As Integer = 8, bizz3 As Integer = 7", "Const bizz1 As Integer = 9, bizz3", "bizz2")]
        [TestCase("Const bizz1 As Integer = 9, bizz2 As Integer = 8, bizz3 As Integer = 7", "Const bizz3 As Integer = 7", "bizz1", "bizz2")]
        [TestCase("Dim bizz1 As Integer, bizz2 As Integer, bizz3 As Integer", "Dim bizz1 As Integer, bizz3", "bizz2")]
        [TestCase("Dim bizz1 As Integer, bizz2 As Integer, bizz3 As Integer", "Dim bizz3 As Integer", "bizz1", "bizz2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeleteDeclarationWithinList(string declaration, string expected, params string[] toDelete)
        {
            var inputCode =
$@"
Public Sub Foo()
    {declaration}
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toDelete));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase("Const const3 As Integer = 7", "const1", "const2")]
        [TestCase("Const const1 As Integer = 9", "const2", "const3")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ConstantsDeclarationListsLineContinuations(string expected, params string[] toDelete)
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, _
            const3 As Integer = 7
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toDelete));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase("Dim bizz3 As Integer", "bizz1", "bizz2")]
        [TestCase("Dim bizz1 As Integer", "bizz2", "bizz3")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariableDeclarationListsLineContinuations(string expected, params string[] toDelete)
        {
            var inputCode =
@"
Public Sub Foo()
    Dim bizz1 As Integer, bizz2 As Integer, _
            bizz3 As Integer
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toDelete));
            StringAssert.Contains(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ConstantsInDeclarationListsLineContinuationsDeleteOneOfTwoInSingleLine_RetainsLineExtension()
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, _
            const3 As Integer = 7
End Sub";

            var expected =
@"
Public Sub Foo()
    Const const1 As Integer = 9, _
            const3 As Integer = 7
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "const2"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariablesInDeclarationListsLineContinuationsDeleteOneOfTwoInSingleLine_RetainsLineExtension()
        {
            var inputCode =
@"
Public Sub Foo()
    Dim bizz1 As Integer, bizz2 As Integer, _
            bizz3 As Integer
End Sub";

            var expected =
@"
Public Sub Foo()
    Dim bizz1 As Integer, _
            bizz3 As Integer
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz2"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Const bizz1 As Integer = 9, bizz2 As Integer = 8, bizz3 As Integer = 7")]
        [TestCase("Dim bizz1 As Integer, bizz2 As Integer, bizz3 As Integer")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeclarationsListsDeleteAll(string declaration)
        {
            var inputCode =
$@"
Public Sub Foo()
    {declaration}
End Sub
";

            var expected =
$@"
Public Sub Foo()
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1", "bizz2", "bizz3"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Dim bizz1 As Long: Dim bizz2 As Long 'Comment on bizz2", "bizz2 As Long 'Comment on bizz2")]
        [TestCase("Dim bizz1 As Long, bizz2 As Long 'Comment on bizz2", "bizz2 As Long 'Comment on bizz2")]
        [TestCase("Const bizz1 As Long = 100: Const bizz2 As Long = 200 'Comment on bizz2", "bizz2 As Long = 200 'Comment on bizz2")]
        [TestCase("Const bizz1 As Long = 100, bizz2 As Long = 200 'Comment on bizz2", "bizz2 As Long = 200 'Comment on bizz2")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeclarationsLineWithTrailingComment_RetainsComments(string declarationList, string expected)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    {declarationList}
    ' More Comments
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.Contains("More Comments", actualCode);
        }

        [TestCase("var1", "var2")]
        [TestCase("var2", "var1")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultiLineVariablesDeclarationWithTrailingComment_RetainsEndOfLineComments(string remove, string retain)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    Dim var1 As String, _
        var2 As String _ 
            'Comment on var1 and or var2
    ' More Comments
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, remove));
            StringAssert.Contains($"Dim {retain} As String", actualCode);
            StringAssert.Contains("'Comment ", actualCode);
            StringAssert.Contains("More Comments", actualCode);
            StringAssert.DoesNotContain($"{remove} As String", actualCode);
        }

        [TestCase("var1", "var2")]
        [TestCase("var2", "var1")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void PartialListDeleteWithTrailingComment_ModifiesComment(string remove, string retain)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    'Preceding Comment
    '@Ignore UseMeaningfulName
    Dim var1 As String, var2 As String 'Comment on var1 and or var2
    ' More Comments
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, remove), (model) => model.InsertValidationTODOForRetainedComments = true);
            StringAssert.Contains($"Dim {retain} As String", actualCode);
            StringAssert.Contains($"{DeleteDeclarationsTestSupport.TodoContent}Preceding Comment", actualCode);
            StringAssert.Contains($"{DeleteDeclarationsTestSupport.TodoContent}Comment on var1 and or var2", actualCode);
            StringAssert.Contains($"{DeleteDeclarationsTestSupport.TodoContent} More Comments", actualCode);
            StringAssert.DoesNotContain($"{remove} As String", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void PartialListDeleteWithAnnotation_RetainsAnnotation() 
        {
            //Annotation is also associated with other declaration(s) in the declaration list.  So, the Annotation cannot be deleted
            //even though the 'other' declaration(s) may not required the Annotation
            var inputCode =
@"
Option Explicit

Sub Foo()
    '@Ignore UseMeaningfulName
    Dim v1 As String, aBetterName As String
End Sub";

            var expected =
@"
Option Explicit

Sub Foo()
    '@Ignore UseMeaningfulName
    Dim aBetterName As String
End Sub";
            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "v1"), (model) => model.InsertValidationTODOForRetainedComments = true);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("bizz1", "bizz2")]
        [TestCase("bizz2", "bizz1")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultiLineConstantsDeclarationWithTrailingComment_RetainsEndOfLineComments(string remove, string retain)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo()
    Const bizz1 As Long = 100, _
        bizz2 As Long = 200 _ 
            'Comment on bizz1 and or bizz2
    ' More Comments
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, remove));
            StringAssert.Contains($"Const {retain} As Long", actualCode);
            StringAssert.Contains("'Comment ", actualCode);
            StringAssert.Contains("More Comments", actualCode);
            StringAssert.DoesNotContain($"Const {remove}", actualCode);
        }

        [TestCase("            Dim varKept As Long", "    'More Comments")]
        [TestCase("            Dim varKept As Long", "    arg = arg * 2")]
        [TestCase("            Dim varKept As Long", "    Dim nextVar As String")]
        [TestCase("            arg = arg + 1", "    'More Comments")]
        [TestCase("            arg = arg + 1", "    arg = arg * 2")]
        [TestCase("            arg = arg + 1", "    Dim nextVar As String")]
        [TestCase("            'Preceeding Comment", "    'More Comments")]
        [TestCase("            'Preceeding Comment", "    arg = arg * 2")]
        [TestCase("            'Preceeding Comment", "    Dim nextVar As String")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariablesWithLineContinuationsWithTrailingAndPreceedingExpressions(string preceeding, string trailing)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo(ByRef arg As Long)
{preceeding}

    Dim bizz1 As Long, _
        bizz2 As Long _
            'Comment on bizz1 and or bizz2


{trailing}
End Sub
";

            var expected =
$@"
Option Explicit

Sub Foo(ByRef arg As Long)
{preceeding}

{trailing}
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1", "bizz2"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("            Dim varKept As Long", "    'More Comments")]
        [TestCase("            Dim varKept As Long", "    arg = arg * 2")]
        [TestCase("            Dim varKept As Long", "    Dim nextVar As String")]
        [TestCase("            arg = arg + 1", "    'More Comments")]
        [TestCase("            arg = arg + 1", "    arg = arg * 2")]
        [TestCase("            arg = arg + 1", "    Dim nextVar As String")]
        [TestCase("            'Preceeding Comment", "    'More Comments")]
        [TestCase("            'Preceeding Comment", "    arg = arg * 2")]
        [TestCase("            'Preceeding Comment", "    Dim nextVar As String")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ConstantsWithLineContinuationsWithTrailingAndPreceedingExpressions(string preceeding, string trailing)
        {
            var inputCode =
$@"
Option Explicit

Sub Foo(ByRef arg As Long)
{preceeding}

    Const bizz1 As Long = 100, _
        bizz2 As Long = 100 _
            'Comment on bizz1 and or bizz2


{trailing}
End Sub
";

            var expected =
$@"
Option Explicit

Sub Foo(ByRef arg As Long)
{preceeding}

{trailing}
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1", "bizz2"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("'@Ignore VariableNotUsed, VariableNotAssigned, UseMeaningfulName", "Dim X As Long")]
        [TestCase("'@Ignore ConstantNotUsed, UseMeaningfulName", "Const X As Long = 7")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void TargetWithAnnotations(string annotations, string declaration)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
    {annotations}
    {declaration}

    Dim usedVar As Long
    usedVar = arg
End Sub
";

            var expected =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
    Dim usedVar As Long
    usedVar = arg
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("'@Ignore VariableNotUsed, VariableNotAssigned, UseMeaningfulName", "Dim X As Long, alsoNotUsed As String")]
        [TestCase("'@Ignore ConstantNotUsed, UseMeaningfulName", "Const X As Long = 7, alsoNotUsed As Long = 9")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void TargetListWithAnnotations(string annotations, string declaration)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
    {annotations}
    {declaration}

    Dim usedVar As Long
    usedVar = arg
End Sub
";

            var expected =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
    Dim usedVar As Long
    usedVar = arg
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X", "alsoNotUsed"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("'@Ignore VariableNotUsed, VariableNotAssigned, UseMeaningfulName", "Dim X As Long, usedVar As Long", "Dim usedVar As Long")]
        [TestCase("'@Ignore ConstantNotUsed, UseMeaningfulName", "Const X As Long = 7, usedVar As Long = 9", "Const usedVar As Long = 9")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void TargetListPartialDeletionWithAnnotations(string annotations, string declaration, string expectedDeclaration)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
    {annotations}
    {declaration}

    arg = arg + usedVar
End Sub
";

            var expected =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
    {annotations}
    {expectedDeclaration}

    arg = arg + usedVar
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("    ", "Dim X As Long", "VariableNotUsed")]
        [TestCase("    ", "Const X As Long = 9", "ConstantNotUsed")]
        [TestCase("        ", "Dim X As Long", "VariableNotUsed")]
        [TestCase("        ", "Const X As Long = 9", "ConstantNotUsed")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleAnnotationLists(string nextStatementIndentation, string declaration, string ignoreNotAssigned)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
    '@Ignore {ignoreNotAssigned}
    '@Ignore UseMeaningfulName
    {declaration}

{nextStatementIndentation}Dim usedVar As Long
    usedVar = 7
    arg = arg + usedVar
End Sub
";

            var expected =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
{nextStatementIndentation}Dim usedVar As Long
    usedVar = 7
    arg = arg + usedVar
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Dim X As Long", "VariableNotUsed")]
        [TestCase("Const X As Long = 9", "ConstantNotUsed")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void AnnotationWithSubsequentComment_RemovesAnnotations(string declaration, string ignoreNotUsed)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    '@Ignore {ignoreNotUsed}, UseMeaningfulName
    'And then another
    {declaration}

    Dim usedVar As Long
    arg = usedVar
End Sub
";

            var expected =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    'And then another
    Dim usedVar As Long
    arg = usedVar
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Dim X As Long", true)]
        [TestCase("Dim X As Long", false)]
        [TestCase("Const X As Long = 9", true)]
        [TestCase("Const X As Long = 9", false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsInjectTODOCommentFlag(string declaration, bool injectTODO)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    '@Ignore UseMeaningfulName
    'And then another
    {declaration}

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";

            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = injectTODO;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            var injectedContent = injectTODO
                ? DeleteDeclarationsTestSupport.TodoContent
                : string.Empty;

            StringAssert.Contains($"{injectedContent}There is already a comment", actualCode);
            StringAssert.Contains($"{injectedContent}And then another", actualCode);
        }

        [TestCase("Dim X As Long", true)]
        [TestCase("Dim X As Long", false)]
        [TestCase("Const X As Long = 9", true)]
        [TestCase("Const X As Long = 9", false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsDeleteAnnotationFlag(string declaration, bool deleteAnnotation)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    '@Ignore UseMeaningfulName
    'And then another
    {declaration}

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";
            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = false;
                model.DeleteAnnotations = deleteAnnotation;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            if (deleteAnnotation)
            {
                StringAssert.DoesNotContain("'@Ignore UseMeaningfulName", actualCode);
            }
            else
            {
                StringAssert.Contains("'@Ignore UseMeaningfulName", actualCode);
            }
        }

        [TestCase("Dim X As Long", true)]
        [TestCase("Dim X As Long", false)]
        [TestCase("Const X As Long = 9", true)]
        [TestCase("Const X As Long = 9", false)]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RespectsDeleteLogicalLineCommentsFlag(string declaration, bool deleteLogicalLineComments)
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    'And then another
    {declaration} 'This is a declaration logical line comment

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";
            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = false;
                model.DeleteDeclarationLogicalLineComments = deleteLogicalLineComments;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            if (deleteLogicalLineComments)
            {
                StringAssert.DoesNotContain("'This is a declaration logical line comment", actualCode);
            }
            else
            {
                StringAssert.Contains("'This is a declaration logical line comment", actualCode);
                StringAssert.DoesNotContain("'And then another 'This is a declaration logical line comment", actualCode);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RetainLogicalLineCommentWithInsertedTODO()
        {
            var inputCode =
$@"
Option Explicit

Public mVar1 As Long

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    'And then another
    Dim X As Long 'This is a declaration logical line comment

    Dim usedVar As Long
    arg = usedVar
    
End Sub
";
            void modelFlags(IDeleteDeclarationsModel model)
            {
                model.InsertValidationTODOForRetainedComments = true;
                model.DeleteDeclarationLogicalLineComments = false;
            }

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"), modelFlags);

            var injectedContent = DeleteDeclarationsTestSupport.TodoContent;

            StringAssert.Contains($"{injectedContent}This is a declaration logical line comment", actualCode);
            StringAssert.DoesNotContain("'{injectedContent}And then another 'This is a declaration logical line comment", actualCode);
        }

        [TestCase("Const bizz As Integer = 9")]
        [TestCase("Dim bizz As Integer")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeleteDeclarationsInMultipleProcedures(string declaration)
        {
            var inputCode =
$@"
Public Sub Foo()
    {declaration}
End Sub

Public Sub Foo2()
    {declaration}
End Sub

Public Sub Foo3()
    {declaration}
End Sub
";

            var expected =
$@"
Public Sub Foo()
End Sub

Public Sub Foo2()
End Sub

Public Sub Foo3()
    {declaration}
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargetsUsingParentDeclaration(state, ("bizz", "Foo"), ("bizz", "Foo2")));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase(" 'Is A DeclarationLogical Line Comment", "'Is A DeclarationLogical Line Comment")]
        [TestCase("'Is A DeclarationLogical Line Comment _\r\n As Is This _\r\n 'As Is This As Well", "Is A DeclarationLogical Line Comment", "As Is This", "'As Is This As Well")]
        [TestCase(" _\r\n    'Is A DeclarationLogical Line Comment", "'Is A DeclarationLogical Line Comment")]
        [TestCase(" _\r\n    'Is A DeclarationLogical Line Comment _\r\n As Is This _\r\n 'As Is This As Well", "Is A DeclarationLogical Line Comment", "As Is This", "'As Is This As Well")]
        [TestCase("")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsDeclarationLogicalLineContext(string declarationLogicalLineContent, params string[] expectedComments)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long)

    Dim X As Long{declarationLogicalLineContent}

End Sub
";
            var expected = declarationLogicalLineContent.Length > 0 ? declarationLogicalLineContent : string.Empty;

            void thisTest(IDeclarationDeletionTarget sut)
            {
                var commentContext = sut.GetDeclarationLogicalLineCommentContext();
                
                var commentContent = commentContext?.GetText() ?? string.Empty;

                if (declarationLogicalLineContent.Length > 0)
                {
                    Assert.IsTrue(commentContext != null);
                    foreach (var expComment in expectedComments)
                    {
                        StringAssert.Contains(expComment, commentContent);
                    }
                }
                else
                {
                    Assert.IsTrue(commentContext is null);
                }
            }

            _support.SetupAndInvokeIDeclarationDeletionTargetTest(inputCode, "X", thisTest);
        }

        [TestCase(" 'Is A DeclarationLogical Line Comment", "'Is A DeclarationLogical Line Comment")]
        [TestCase("'Is A DeclarationLogical Line Comment _\r\n As Is This _\r\n 'As Is This As Well", "Is A DeclarationLogical Line Comment", "As Is This", "'As Is This As Well")]
        [TestCase(" _\r\n    'Is A DeclarationLogical Line Comment", "'Is A DeclarationLogical Line Comment")]
        [TestCase(" _\r\n    'Is A DeclarationLogical Line Comment _\r\n As Is This _\r\n 'As Is This As Well", "Is A DeclarationLogical Line Comment", "As Is This", "'As Is This As Well")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void FindsDeclarationLogicalLineContext_ListDeclaration(string declarationLogicalLineContent, params string[] expectedComments)
        {
            var inputCode =
$@"
Option Explicit

Public Const notUsed1 As Long = 100, _
    notUsed2 As Long = 200, _
        notUsed3 As Long = 300 _
            {declarationLogicalLineContent}


'This field is used
Public used As String
";

            void thisTest(IDeclarationDeletionTarget sut)
            {
                var commentContext = sut.GetDeclarationLogicalLineCommentContext();

                var commentContent = commentContext?.GetText() ?? string.Empty;

                if (declarationLogicalLineContent.Length > 0)
                {
                    Assert.IsTrue(commentContext != null);
                    foreach (var expComment in expectedComments)
                    {
                        StringAssert.Contains(expComment, commentContent);
                    }
                }
                else
                {
                    Assert.IsFalse(commentContext is null);
                }
            }

            _support.SetupAndInvokeIDeclarationDeletionTargetTest(inputCode, "notUsed1", thisTest);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void EmbeddedAnotationAndCommentWhileWendStmt()
        {
            var inputCode =
@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim i As Long: i = 0
    While i < 10 'Add a sub collection
        '@Ignore UseMeaningfulName
        Dim sC As Collection 'This collection is added to the mainCollection
        Set sC = New Collection
        mainCollection.Add sC
    Wend
End Sub
";

            var expected =
@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim i As Long: i = 0
    While i < 10 'Add a sub collection
        Set sC = New Collection
        mainCollection.Add sC
    Wend
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "sC"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Do While i < 10", "Loop")]
        [TestCase("Do Until i >= 10", "Loop")]
        [TestCase("Do", "Loop While i < 10")]
        [TestCase("Do", "Loop Until i >= 10")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void EmbeddedAnotationAndCommentDoLoops(string loopStart, string loopEnd)
        {
            var inputCode =
$@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim i As Long: i = 0
    {loopStart} 'Add a sub collection
        '@Ignore UseMeaningfulName
        Dim sC As Collection 'This collection is added to the mainCollection
        Set sC = New Collection
        mainCollection.Add sC
        i = i + 1
    {loopEnd}
End Sub
";

            var expected =
$@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim i As Long: i = 0
    {loopStart} 'Add a sub collection
        Set sC = New Collection
        mainCollection.Add sC
        i = i + 1
    {loopEnd}
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "sC"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void EmbeddedAnotationAndCommentWithStatement()
        {
            var inputCode =
@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    With mainCollection 'Add a sub collection
        Dim sC As Collection 'This collection is added to the mainCollection
        Set sC = New Collection
        .Add sC
    End With
End Sub
";

            var expected =
@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    With mainCollection 'Add a sub collection
        Set sC = New Collection
        .Add sC
    End With
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "sC"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }
    }
}
