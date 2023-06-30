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
    public class DeleteDeclarationsLocalsUnterminatedBlockContextTests : DeleteDeclarationsLocalsTestsBase
    {
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void EmbeddedAnotationAndCommentForNextStmt()
        {
            var inputCode =
@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim i As Long
    For i = 1 To 10 'Add a sub collection
        '@Ignore UseMeaningfulName
        Dim sC As Collection 'This collection is added to the mainCollection
        Set sC = New Collection
        mainCollection.Add sC
    Next i
End Sub
";

            var expected =
@"
Public Sub Test()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim i As Long
    For i = 1 To 10 'Add a sub collection
        Set sC = New Collection
        mainCollection.Add sC
    Next i
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "sC"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void EmbeddedAnotationAndCommentForEachStmt()
        {
            var inputCode =
@"
Public Sub TestForEach()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim letters As Collection
    Set letters = New Collection
    letters.Add ""A""
    letters.Add ""B""
    letters.Add ""C""

    Dim letter As Variant
    For Each letter In letters  'Add a string
        '@Ignore UseMeaningfulName
        Dim xC As String 'This string fragment is appended to each letter
        xC = ""_Suffix""
        mainCollection.Add letter & xC
    Next
End Sub
";

            var expected =
@"
Public Sub TestForEach()
    Dim mainCollection As Collection
    Set mainCollection = New Collection
    
    Dim letters As Collection
    Set letters = New Collection
    letters.Add ""A""
    letters.Add ""B""
    letters.Add ""C""

    Dim letter As Variant
    For Each letter In letters  'Add a string
        xC = ""_Suffix""
        mainCollection.Add letter & xC
    Next
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "xC"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }
    }
}
