using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.DeleteDeclarations.DeleteProcedureScopeElements
{
    [TestFixture]
    public class DeleteDeclarationsLocalMultipleDeleteGroupsTests : DeleteDeclarationsLocalsTestsBase
    {
        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeleteGroups()
        {
            var inputCode =
@"
Public Sub Test()
    Dim firstLong As Long 'Group1

    Dim mainCollection As Collection

    Dim firstStr As String 'Group2
    Dim secondStr As String 'Group2
    
    Dim thirdStr As String 'Group2

    Dim firstVar As Variant

    Dim i As Long

    Dim firstBool As Boolean

    Dim secondBool As Boolean 'Group3
End Sub
";

            var expected =
@"
Public Sub Test()
    Dim mainCollection As Collection

    Dim firstVar As Variant

    Dim i As Long

    Dim firstBool As Boolean

End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "firstLong", "firstStr", "secondStr", "thirdStr", "secondBool"));
            StringAssert.Contains(expected, actualCode);
            StringAssert.AreEqualIgnoringCase(expected, actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeleteGroupsForNext()
        {
            var inputCode =
@"
Public Sub Test()
    Dim firstLong As Long 'Group1

    Dim mainCollection As Collection

    Dim i As Long
    For i = 0 To 10
        Dim firstStr As String 'Group2
        Dim secondStr As String 'Group2
    
        Dim thirdStr As String 'Group2
    Next i

    Dim firstVar As Variant

    Dim firstBool As Boolean

    Dim secondBool As Boolean 'Group3
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "firstLong", "firstStr", "secondStr", "thirdStr", "secondBool"));
            StringAssert.Contains("\r\n    Next i", actualCode);
            StringAssert.DoesNotContain("Dim firstStr", actualCode);
            StringAssert.DoesNotContain("Dim secondtStr", actualCode);
            StringAssert.DoesNotContain("Dim thirdStr", actualCode);
            StringAssert.DoesNotContain("'Group2", actualCode);
        }

        [Test]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void MultipleDeleteGroupsForNextNested()
        {
            var inputCode =
@"
Public Sub Test()
    Dim firstLong As Long 'Group1

    Dim mainCollection As Collection

    Dim i As Long
    For i = 0 To 10
        Dim firstStr As String 'Group2

        For i = 0 To 10
            Dim secondStr As String 'Group3
    
            Dim thirdStr As String 'Group3
        Next i
    Next i

    Dim firstVar As Variant

    Dim firstBool As Boolean

    Dim secondBool As Boolean 'Group4
End Sub
";

            var actualCode = GetRetainedCodeBlock(inputCode, state => _support.TestTargets(state, "firstLong", "firstStr", "secondStr", "thirdStr", "secondBool"));
            StringAssert.Contains("\r\n    Next i", actualCode);
            StringAssert.Contains("\r\n        Next i", actualCode);
            StringAssert.DoesNotContain("Dim firstStr", actualCode);
            StringAssert.DoesNotContain("Dim secondtStr", actualCode);
            StringAssert.DoesNotContain("Dim thirdStr", actualCode);
            StringAssert.DoesNotContain("'Group2", actualCode);
            StringAssert.DoesNotContain("'Group3", actualCode);
        }
    }
}
