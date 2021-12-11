using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class DeleteDeclarationsLocalsWithLabelsTests : DeleteDeclarationsLocalsTestsBase
    {
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, targets));
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "target"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Dim target As Long, bogey As String, bogey2 As Integer", "Dim bogey As String, bogey2 As Integer")]
        [TestCase(@"Const target As Long = 100, bogey As String = ""Yo!!"", bogey2 As Integer = 5", @"Const bogey As String = ""Yo!!"", bogey2 As Integer = 5")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveDeclarationFollowingLabel_PartialRemoval(string expression, string expectedExpression)
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

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "target"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [TestCase("Dim target As Long, bogey As String, bogey2 As Integer")]
        [TestCase(@"Const target As Long = 100, bogey As String = ""Yo!!"", bogey2 As Integer = 5")]
        [Category("Refactorings")]
        [Category("DeleteDeclarationWithLineLabel")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void RemoveDeclarationFollowingLabel_FullRemoval(string expression)
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

Label1:    

    Dim var2 As Variant
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "target", "bogey", "bogey2"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
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

            var expected =
                $@"Private Sub Foo()
1{indent}
2{indent}Dim bat As Integer
3{indent}bizz = True
End Sub";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "bizz"));
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithPrecedingAnnotationEndStatementColon()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

    Dim var2 As Variant

'@Ignore UseMeaningfulName
Label1:    Dim X As Long: X = arg

End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.Contains("'@Ignore UseMeaningfulName", actualCode);
            StringAssert.DoesNotContain("Dim X As Long", actualCode);
            StringAssert.DoesNotContain(": X", actualCode);
            StringAssert.Contains("Dim var2 As Variant", actualCode);
            StringAssert.Contains("Label1:", actualCode);
            StringAssert.Contains("X = arg", actualCode);
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

            var expected =
$@"
Option Explicit

Public Sub DoSomething(ByRef arg As Long)
    '@Ignore {ignoreNotUsed}
    '@Ignore UseMeaningfulName
1   

    Dim usedVar As Long
    usedVar = 7
    arg = arg + usedVar
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.Contains($"'@Ignore {ignoreNotUsed}", actualCode);
            StringAssert.Contains("'@Ignore UseMeaningfulName", actualCode);
            StringAssert.DoesNotContain($"{declaration}", actualCode);
            StringAssert.Contains("1 ", actualCode);
            StringAssert.Contains("Dim usedVar As Long", actualCode);
            StringAssert.Contains("usedVar = 7", actualCode);
            StringAssert.Contains("arg = arg + usedVar", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelWithPrecedingAnnotation()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

    Dim var2 As Variant

'@Ignore UseMeaningfulName
Label1:    Dim X As Long
    X = arg

End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X"));
            StringAssert.Contains("'@Ignore UseMeaningfulName", actualCode);
            StringAssert.DoesNotContain("Dim X As Long", actualCode);
            StringAssert.Contains("Dim var2 As Variant", actualCode);
            StringAssert.Contains("Label1:", actualCode);
            StringAssert.Contains("X = arg", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelsSameLogicalLineDeletedAndRetained_GetsCorrectSpacing()
        {
            const string inputCode =
@"
Sub Foo(ByVal arg As Long)

    Dim var2 As Variant

Label1:    Dim X As Long
    X = arg

Label2:    Dim Y As String

End Sub
";
            var expected =
@"
Sub Foo(ByVal arg As Long)

    Dim var2 As Variant

Label1:    
    X = arg

End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "X", "Y", "Label2"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }
    }
}
