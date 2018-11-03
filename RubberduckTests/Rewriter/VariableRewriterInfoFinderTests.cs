using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Rewriter.RewriterInfo;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class VariableRewriterInfoFinderTests
    {

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode = 
                @"Dim fooBar As Long
    ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_UnterminatedBlock()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim col As Collection
    Dim bar As Variant
    For Each bar in col
        Dim fooBar As Long
        Baz bar
    Next
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long
        ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_StatementSeparator()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long : Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long : ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_Comment()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long 'Comment
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_NextLineHasLabel()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long
hm: Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_NextLineHasLineNumber()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long
1   Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_FollowedByMultipleEndOfStatements_EndOfLineFirst()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long
    
    
    :
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long
    ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleLocalVariable_FollowedByMultipleEndOfStatements_StatementSeparatorFirst()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long :
    
    
    :
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Dim fooBar As Long :";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_LocalVariableInList_First()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBar As Long, fooBaz As Long, barBaz As Long
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long, ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_LocalVariableInList_Middle()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim fooBaz As Long, fooBar As Long, barBaz As Long
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long, ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_LocalVariableInList_Last()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Dim barBaz As Long, fooBaz As Long, fooBar As Long
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @", fooBar As Long";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_ModuleVariableInList_First()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBar As Long, fooBaz As Long, barBaz As Long
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long, ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_ModuleVariableInList_Middle()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBaz As Long, fooBar As Long, barBaz As Long
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long, ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_ModuleVariableInList_Last()
        {
            var inputCode =
                @"
Public bar As Integer
Private barBaz As Long, fooBaz As Long, fooBar As Long
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @", fooBar As Long";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleModuleVariable()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBar As Long
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Private fooBar As Long
";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleModuleVariable_StatementSeparator()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBar As Long : Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Private fooBar As Long : ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleModuleVariable_Comment()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBar As Long 'Comment
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Private fooBar As Long ";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleModuleVariable_FollowedByMultipleEndOfStatements_EndOfLineFirst()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBar As Long


    :     
 :


Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Private fooBar As Long
";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestVariableRewriterInfoFinder_SingleModuleVariable_FollowedByMultipleEndOfStatements_StatementSeparatorFirst()
        {
            var inputCode =
                @"
Public bar As Integer
Private fooBar As Long :


    :     
 :


Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var variableName = "fooBar";
            var expectedEnclosedCode =
                @"Private fooBar As Long :";

            TestEnclosedCode(inputCode, variableName, expectedEnclosedCode);
        }

        private void TestEnclosedCode(string inputCode, string variableName, string expectedEnclosedCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var variableSubStmtContext = state.DeclarationFinder.MatchName(variableName).First().Context;

                var infoFinder = new VariableRewriterInfoFinder();
                var info = infoFinder.GetRewriterInfo(variableSubStmtContext);

                var actualEnclosedCode = state.GetCodePaneTokenStream(component.QualifiedModuleName)
                    .GetText(info.StartTokenIndex, info.StopTokenIndex);
                Assert.AreEqual(expectedEnclosedCode, actualEnclosedCode);
            }
        }
    }
}
