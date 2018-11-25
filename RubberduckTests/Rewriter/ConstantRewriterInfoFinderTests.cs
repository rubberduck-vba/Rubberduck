using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Rewriter.RewriterInfo;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class ConstantRewriterInfoFinderTests
    {

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1
    ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_UnterminatedBlock()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim col As Collection
    Dim bar As Variant
    For Each bar in col
        Const fooBar As Long = 1
        Baz bar
    Next
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1
        ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_StatementSeparator()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1 : Baz bar 
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1 : ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_Comment()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1 'Comment
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1 ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_NextLineHasLabel()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1
hm: Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_NextLineHasLineNumber()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1
1   Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_FollowedByMultipleEndOfStatements_EndOfLineFirst()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1
    
    
    :
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1
    ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleLocalConstant_FollowedByMultipleEndOfStatements_StatementSeparatorFirst()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1 :
    
    
    :
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Const fooBar As Long = 1 :";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_LocalConstantInList_First()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBar As Long = 1, fooBaz As Long = 2, barBaz As Long = 3
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long = 1, ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_LocalConstantInList_Middle()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const fooBaz As Long = 1, fooBar As Long = 2, barBaz As Long = 3
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long = 2, ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_LocalConstantInList_Last()
        {
            var inputCode =
                @"Public Sub Foo()
    Dim bar As Integer
    Const barBaz As Long = 1, fooBaz As Long = 2, fooBar As Long = 3
    Baz bar
End Sub

Private Sub Baz(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @", fooBar As Long = 3";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_ModuleConstantInList_First()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBar As Long = 1, fooBaz As Long = 2, barBaz As Long = 3
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long = 1, ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_ModuleConstantInList_Middle()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBaz As Long = 1, fooBar As Long = 2, barBaz As Long = 3
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"fooBar As Long = 2, ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_ModuleConstantInList_Last()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const barBaz As Long = 1, fooBaz As Long = 2, fooBar As Long = 3
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @", fooBar As Long = 3";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleModuleConstant()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBar As Long = 1
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Private Const fooBar As Long = 1
";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleModuleConstant_StatementSeparator()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBar As Long = 1 : Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Private Const fooBar As Long = 1 : ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleModuleConstant_Comment()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBar As Long = 1 'Comment
Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Private Const fooBar As Long = 1 ";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleModuleConstant_FollowedByMultipleEndOfStatements_EndOfLineFirst()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBar As Long = 1


    :     
 :


Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Private Const fooBar As Long = 1
";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        [Test]
        [Category("Rewriter")]
        public void TestConstantRewriterInfoFinder_SingleModuleConstant_FollowedByMultipleEndOfStatements_StatementSeparatorFirst()
        {
            var inputCode =
                @"
Public bar As Integer
Private Const fooBar As Long = 1 :


    :     
 :


Private baz As String

Private Sub Whatever(brr As Integer)
End Sub";
            var constantName = "fooBar";
            var expectedEnclosedCode =
                @"Private Const fooBar As Long = 1 :";

            TestEnclosedCode(inputCode, constantName, expectedEnclosedCode);
        }

        private void TestEnclosedCode(string inputCode, string constantName, string expectedEnclosedCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var constantSubStmtContext = state.DeclarationFinder.MatchName(constantName).First().Context;

                var infoFinder = new ConstantRewriterInfoFinder();
                var info = infoFinder.GetRewriterInfo(constantSubStmtContext);

                var actualEnclosedCode = state.GetCodePaneTokenStream(component.QualifiedModuleName)
                    .GetText(info.StartTokenIndex, info.StopTokenIndex);
                Assert.AreEqual(expectedEnclosedCode, actualEnclosedCode);
            }
        }
    }
}
