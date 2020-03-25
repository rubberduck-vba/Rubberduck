using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnusedDeclarationQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ConstantNotUsed_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ConstantNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void LabelNotUsed_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
label1:
End Sub";

            const string expectedCode =
                @"Sub Foo()

End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new LineLabelNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void LabelNotUsed_QuickFixWorks_MultipleLabels()
        {
            const string inputCode =
                @"Sub Foo()
label1:
dim var1 as variant
label2:
goto label1:
End Sub";

            const string expectedCode =
                @"Sub Foo()
label1:
dim var1 as variant

goto label1:
End Sub"; ;

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new LineLabelNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureNotUsed_QuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode = @"";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String
End Sub";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_WithFollowingEmptyLine_DoesNotRemoveEmptyLine()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String

End Sub";

            const string expectedCode =
                @"Sub Foo()

End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_WithCommentOnSameLine_DoesNotRemoveComment()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String ' Comment
End Sub";

            const string expectedCode =
                @"Sub Foo()
' Comment
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_WithCommentOnSameLineAndFollowingStuff_DoesNotRemoveComment()
        {
            const string inputCode =
                @"Function Foo() As String
Dim var1 As String ' Comment
Dim var2 As String
var2 = ""Something""
Foo = var2
End Function";

            const string expectedCode =
                @"Function Foo() As String
' Comment
Dim var2 As String
var2 = ""Something""
Foo = var2
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }



        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_WithFollowingCommentLine_DoesNotRemoveCommentLine()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String
' Comment
End Sub";

            const string expectedCode =
                @"Sub Foo()
' Comment
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_InMultideclaration_WithFollowingCommentLine_DoesNotRemoveCommentLineOrOtherDeclarations()
        {
            const string inputCode =
                @"Function Foo() As String
Dim var1 As String, var2 As String
' Comment
var2 = ""Something""
Foo = var2
End Function";

            const string expectedCode =
                @"Function Foo() As String
Dim var2 As String
' Comment
var2 = ""Something""
Foo = var2
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_InMultideclarationByStmtSeparators_WithFollowingCommentLine_DoesNotRemoveCommentLineOrOtherDeclarations()
        {
            const string inputCode =
                @"Function Foo() As String
Dim var1 As String:Dim var2 As String
' Comment
var2 = ""Something""
Foo = var2
End Function";

            const string expectedCode =
                @"Function Foo() As String
Dim var2 As String
' Comment
var2 = ""Something""
Foo = var2
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveUnusedDeclarationQuickFix();
        }
    }
}
