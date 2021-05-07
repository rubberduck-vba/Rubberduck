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
    public class DeclarationDeleter_LocalVariablesAndConstantsTests
    {
        private readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

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

            var actualCodeLines = _support.GetRetainedLines(inputCode, state => _support.TestTargets(state, "bizz"));
            Assert.IsFalse(actualCodeLines.Contains(declaration));
            Assert.AreEqual(2, actualCodeLines.Count());
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

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toDelete));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase("Const const3 As Integer = 7", "const1", "const2")]
        [TestCase("Const const1 As Integer = 9", "const2", "const3")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ConstantsDeclarationListsLineContinuations(string expectedCode, params string[] toDelete)
        {
            var inputCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, const2 As Integer = 8, _
            const3 As Integer = 7
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toDelete));
            StringAssert.Contains(expectedCode, actualCode);
        }

        [TestCase("Dim bizz3 As Integer", "bizz1", "bizz2")]
        [TestCase("Dim bizz1 As Integer", "bizz2", "bizz3")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void VariableDeclarationListsLineContinuations(string expectedCode, params string[] toDelete)
        {
            var inputCode =
@"
Public Sub Foo()
    Dim bizz1 As Integer, bizz2 As Integer, _
            bizz3 As Integer
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, toDelete));
            StringAssert.Contains(expectedCode, actualCode);
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

            var expectedCode =
@"
Public Sub Foo()
    Const const1 As Integer = 9, _
            const3 As Integer = 7
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "const2"));
            StringAssert.Contains(expectedCode, actualCode);
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

            var expectedCode =
@"
Public Sub Foo()
    Dim bizz1 As Integer, _
            bizz3 As Integer
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz2"));
            StringAssert.Contains(expectedCode, actualCode);
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
End Sub";

            var actualCodeLines = _support.GetRetainedLines(inputCode, state => _support.TestTargets(state, "bizz1", "bizz2", "bizz3"));
            Assert.IsTrue(2 == actualCodeLines.Count(), $"Unexpected line count: {Environment.NewLine} {string.Join(Environment.NewLine, actualCodeLines)}");
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
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1"));
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

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, remove));
            StringAssert.Contains($"Dim {retain} As String", actualCode);
            StringAssert.Contains("'Comment ", actualCode);
            StringAssert.Contains("More Comments", actualCode);
            StringAssert.DoesNotContain($"Dim {remove}", actualCode);
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

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, remove));
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

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1", "bizz2"));
            StringAssert.Contains(expected, actualCode);
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

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz1", "bizz2"));
            StringAssert.Contains(expected, actualCode);
        }

        [TestCase("Dim bar As Boolean", "1", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String", "1   Dim bazz As String", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String, bizz As String", "1   Dim bar As Boolean", "bizz", "bazz")]
        [TestCase("Const bar As Long = 100", "1", "bar")]
        [TestCase("Const bar As Long = 100, bazz As Long = 200", "1   Const bazz As Long = 200", "bar")]
        [TestCase("Const bar As Long = 100, bazz As Long = 200, bizz As Long = 300", "1   Const bar As Long = 100", "bizz", "bazz")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeletionTargetWithPreceedingLineNumber(string expression, string expected, params string[] targets)
        {
            var inputCode =
$@"
Private Sub Foo()
1   {expression}
2   Dim bat As Integer
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.Contains("2   Dim bat As Integer", actualCode);
            foreach (var deletedIdentifier in targets)
            {
                StringAssert.DoesNotContain(deletedIdentifier, actualCode);
            }
        }

        [TestCase("Dim bar As Boolean", "Label1:", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String", "Label1:   Dim bazz As String", "bar")]
        [TestCase("Dim bar As Boolean, bazz As String, bizz As String", "Label1:   Dim bar As Boolean", "bizz", "bazz")]
        [TestCase("Const bar As Long = 100", "Label1:", "bar")]
        [TestCase("Const bar As Long = 100, bazz As Long = 200", "Label1:   Const bazz As Long = 200", "bar")]
        [TestCase("Const bar As Long = 100, bazz As Long = 200, bizz As Long = 300", "Label1:   Const bar As Long = 100", "bizz", "bazz")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void DeletionTargetWithPreceedingLineLabel(string expression, string expected, params string[] targets)
        {
            var inputCode =
$@"
Private Sub Foo()
Label1:   {expression}
   Dim bat As Integer
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
            StringAssert.Contains(expected, actualCode);
            StringAssert.Contains("   Dim bat As Integer", actualCode);
            foreach (var deletedIdentifier in targets)
            {
                StringAssert.DoesNotContain(deletedIdentifier, actualCode);
            }
        }
        [TestCase("Dim target As Long: target = arg", "target = arg")]
        [TestCase("Const target As Long = 100: arg = target * arg", "arg = target * arg")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveDeclarationFollowingLabel(string expression, string expectedExpression)
        {
            var inputCode =
$@"
Sub Foo(ByRef arg As Long)

Label1:    {expression}

    Dim var2 As Variant
End Sub";

            var expected =
$@"
Sub Foo(ByRef arg As Long)

Label1:    {expectedExpression}

    Dim var2 As Variant
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "target"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.Contains("Dim var2", actualCode);
        }

        [TestCase("Dim bizz As Boolean")]
        [TestCase("Const bizz As Boolean = True")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void TargetWithPreceedingLineNumber(string declaration)
        {
            var indent = "   ";

            var inputCode =
                $@"Private Sub Foo()
1{indent}{declaration}
2{indent}Dim bat As Integer
3{indent}bizz = True
End Sub";

            var expectedCode =
                $@"Private Sub Foo()
1{indent}
2{indent}Dim bat As Integer
3{indent}bizz = True
End Sub";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz"));
            StringAssert.Contains(expectedCode, actualCode);
            StringAssert.AreEqualIgnoringCase(expectedCode, actualCode);
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

            var expectedCode =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
    Dim usedVar As Long
    usedVar = arg
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            Assert.AreEqual(expectedCode, actualCode);
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

            var expectedCode =
$@"
Option Explicit

Public Sub DoSomething(arg As Long)
    Dim usedVar As Long
    usedVar = arg
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X", "alsoNotUsed"));
            StringAssert.Contains(expectedCode, actualCode);
            StringAssert.AreEqualIgnoringCase(expectedCode, actualCode);
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

            var expectedCode =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
    {annotations}
    {expectedDeclaration}

    arg = arg + usedVar
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            Assert.AreEqual(expectedCode, actualCode);
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

            var expectedCode =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
{nextStatementIndentation}Dim usedVar As Long
    usedVar = 7
    arg = arg + usedVar
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.Contains(expectedCode, actualCode);
        }

        [TestCase("Dim X As Long", "VariableNotUsed")]
        [TestCase("Const X As Long = 9", "ConstantNotUsed")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void AnnotationPrecedingLineNumberLabelAndDeclarationLine(string declaration, string ignoreNotUsed)
        {
            var inputCode =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
    '@Ignore {ignoreNotUsed}
    '@Ignore UseMeaningfulName
1   {declaration}

    Dim usedVar As Long
    usedVar = 7
    arg = arg + usedVar
End Sub
";

            var expectedCode =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
1   

    Dim usedVar As Long
    usedVar = 7
    arg = arg + usedVar
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.Contains(expectedCode, actualCode);
            Assert.AreEqual(expectedCode, actualCode);
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

            var expectedCode =
$@"
Option Explicit

Public Sub DoSomethingElse(arg As Long)
    'There is already a comment
    'And then another
    Dim usedVar As Long
    arg = usedVar
End Sub
";

            var actualCode = _support.GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.Contains(expectedCode, actualCode);
            Assert.AreEqual(expectedCode, actualCode);
        }
    }
}
