using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnusedDeclarationQuickFixTests
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ConstantNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new RemoveUnusedDeclarationQuickFix(state).Fix(inspectionResults.First());

                var rewriter = state.GetRewriter(component);
                var rewrittenCode = rewriter.GetText();
                Assert.AreEqual(expectedCode, rewrittenCode);
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureNotUsed_QuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode = @"";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new RemoveUnusedDeclarationQuickFix(state).Fix(inspectionResults.First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

                var rewriter = state.GetRewriter(component);
                Assert.AreEqual(expectedCode, rewriter.GetText());
            }
        }

    }
}
