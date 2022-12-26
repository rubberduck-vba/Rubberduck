using NUnit.Framework;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestFixture]
    public class ExtractMethodSelectionValidationTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        [TestCase(4, 4, 10, 14, Description = "Start in one sub and end in another function")]
        [TestCase(2, 1, 2, 16, Description = "Option Explicit statement")]
        [TestCase(3, 1, 6, 8, Description = "Entire sub so no enclosing method")]
        [TestCase(4, 5, 4, 10, Description = "Partial declaration")]
        public void ValidationShouldReturnFalse(int startLine, int startColumn, int endLine, int endColumn)
        {
            var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub


Private Function NewMethod
    dim a as string
    Debug.Print a
End Function


Private Sub NewMethod4
    dim a as string

    Debug.Print a
End Sub";

            using (var state = MockParser.ParseString(inputCode, out QualifiedModuleName qualifiedModuleName))
            {
                var declarations = state.AllDeclarations;
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var SUT = new ExtractMethodSelectionValidation(declarations, state.ProjectsProvider);

                var actual = SUT.IsSelectionValid(qSelection.Value);
                var expected = false;
                Assert.AreEqual(expected, actual);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        [TestCase(4, 1, 4, 21, Description = "Just a declaration including whitespace")]
        [TestCase(4, 1, 5, 14, Description = "All contents of a sub")]
        [TestCase(11, 1, 11, 18, Description = "Debug.Print statement")]
        public void ValidationShouldReturnTrue(int startLine, int startColumn, int endLine, int endColumn)
        {
            var inputCode = @"
Option Explicit
Private Sub Foo()
    Dim x As Integer
    x = 1 + 2
End Sub


Private Function NewMethod
    dim a as string
    Debug.Print a
End Function


Private Sub NewMethod4
    dim a as string

    Debug.Print a
End Sub";

            using (var state = MockParser.ParseString(inputCode, out QualifiedModuleName qualifiedModuleName))
            {
                var declarations = state.AllDeclarations;
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var SUT = new ExtractMethodSelectionValidation(declarations, state.ProjectsProvider);

                var actual = SUT.IsSelectionValid(qSelection.Value);
                var expected = true;
                Assert.AreEqual(expected, actual);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        [TestCase(3, 1, 7, 8, Description = "Entire compiler directive section")]
        [TestCase(3, 1, 4, 22, Description = "Partial compiler directive section")]
        public void ValidationFailsForCompilerDirectives(int startLine, int startColumn, int endLine, int endColumn)
        {
            var inputCode = @"
Private Sub Foo()
#If VBA7 = 1 Then
    Dim i As LongLong
#Else
    Dim i As Long
#End If
    i = 1 + 2
    Debug.Print i
End Sub";

            using (var state = MockParser.ParseString(inputCode, out QualifiedModuleName qualifiedModuleName))
            {
                var declarations = state.AllDeclarations;
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var SUT = new ExtractMethodSelectionValidation(declarations, state.ProjectsProvider);

                var isValid = SUT.IsSelectionValid(qSelection.Value);
                var actual = SUT.ContainsCompilerDirectives;
                var expected = true;
                Assert.AreEqual(expected, actual);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("ExtractMethod")]
        [TestCase(4, 1, 4, 26, Description = "Invalid if On Error")]
        [TestCase(5, 1, 5, 13, Description = "Invalid if raise error")]
        [TestCase(6, 1, 6, 8, Description = "Invalid if End statement")]
        [TestCase(7, 1, 7, 13, Description = "Invalid if GoSub statement")]
        [TestCase(8, 1, 8, 13, Description = "Invalid if Exit statement")]
        [TestCase(9, 1, 9, 4, Description = "Invalid if line label statement")]
        [TestCase(10, 1, 10, 11, Description = "Invalid if Return statement")]
        [TestCase(14, 1, 14, 11, Description = "Invalid if Resume statement")]
        [TestCase(15, 1, 15, 25, Description = "Invalid if On ... GoSub statement")]
        [TestCase(16, 1, 16, 24, Description = "Invalid if On ... GoTo statement")]
        public void ValidationFailsForUnsupportedCode(int startLine, int startColumn, int endLine, int endColumn)
        {
            var inputCode = @"
Private Sub Foo()
    Dim i As Long
    On Error GoTo exitsub
    Error 11
    End
    GoSub 20
    Exit Sub
20:
    Return
    i = 1 + 2
    Debug.Print i
exitsub:
    Resume
    On 2 GoSub 20, final
    On 2 GoTo 20, final
final:
End Sub";

            using (var state = MockParser.ParseString(inputCode, out QualifiedModuleName qualifiedModuleName))
            {
                var declarations = state.AllDeclarations;
                var selection = new Selection(startLine, startColumn, endLine, endColumn);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

                var SUT = new ExtractMethodSelectionValidation(declarations, state.ProjectsProvider);

                var actual = SUT.IsSelectionValid(qSelection.Value);
                var expected = false;
                Assert.AreEqual(expected, actual);
            }
        }
    }
}