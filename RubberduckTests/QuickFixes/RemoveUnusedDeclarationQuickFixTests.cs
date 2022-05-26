using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using Rubberduck.SmartIndenter;
using RubberduckTests.Mocks;
using RubberduckTests.Refactoring.DeleteDeclarations;
using RubberduckTests.Settings;
using System;
using System.Linq;
using System.Threading;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnusedDeclarationQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
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
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
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
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void LabelNotUsed_QuickFixWorks_MultipleLabels()
        {
            const string inputCode =
                @"Sub Foo()
label1:
Dim var1 As Variant
label2:
goto label1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new LineLabelNotUsedInspection(state));
            var lines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Select(l => l.Trim());
            Assert.IsTrue(lines.Contains("Dim var1 As Variant"));
            Assert.IsTrue(lines.Contains("goto label1"));
            Assert.IsFalse(lines.Contains("label2"));
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
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
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
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
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
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
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UnassignedVariable_WithCommentOnSameLine_DoesNotRemoveComment()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String ' Comment
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            var lines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Select(l => l.Trim());
            Assert.IsFalse(lines.Contains("' Comment"));
            Assert.IsFalse(lines.Contains("var1"));
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UnassignedVariable_WithCommentOnSameLineAndFollowingStuff_DoesNotRemoveComment()
        {
            const string inputCode =
                @"Function Foo() As String
Dim var1 As String ' Comment
Dim var2 As String
var2 = ""Something""
Foo = var2
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            var lines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringAssert.AreEqualIgnoringCase("Dim var2 As String", lines[1].Trim());
            StringAssert.DoesNotContain("var1", actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UnassignedVariable_WithFollowingCommentLine_DoesNotRemoveCommentLine()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String
    ' Comment
Dim var2 As String
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            var lines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Select(l => l.Trim());
            Assert.IsTrue(lines.Contains($"{DeleteDeclarationsTestSupport.TodoContent} Comment"));
            Assert.IsFalse(lines.Contains("var1"));
        }

        [TestCase("Dim var1 As String, var2 As String")]
        [TestCase("Dim var1 As String: Dim var2 As String")]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UnassignedLocalVariable_InMultideclaration_WithFollowingCommentLine_DoesNotRemoveCommentLineOrOtherDeclarations(string declarations)
        {
            var inputCode =
                $@"Function Foo() As String
{declarations}
' Comment
var2 = ""Something""
Foo = var2
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotUsedInspection(state));
            var lines = actualCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringAssert.AreEqualIgnoringCase("Dim var2 As String", lines[1].Trim());
            StringAssert.Contains(" Comment", lines[2].Trim());
            StringAssert.DoesNotContain("var1", actualCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5719
        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UnReferencedProperty_RemovesExtraNewLines()
        {
            const string inputCode =
@"
Public Property Get ReadOnlyProperty() As String
End Property
Public Property Let ReadOnlyProperty(ByVal RHS As String)
End Property

Public Property Get SomeOtherProperty() As String
End Property

Private Sub ReferencingSub()
   Dim tester As String
   tester = ReadOnlyProperty
End Sub
";

            const string expectedCode =
@"
Public Property Get ReadOnlyProperty() As String
End Property

Public Property Get SomeOtherProperty() As String
End Property

Private Sub ReferencingSub()
   Dim tester As String
   tester = ReadOnlyProperty
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureNotUsedInspection(state));
            StringAssert.StartsWith(expectedCode.Trim(), actualCode.Trim());
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void UnReferencedProperties_RemovesExtraNewLines()
        {
            const string inputCode =
@"
'Only calling 'ReferencingSub' here to give it a reference for the test
Public Property Get ReadOnlyProperty() As String
    ReferencingSub ""arg""
End Property

Public Property Let ReadOnlyProperty(ByVal RHS As String)
End Property

Public Property Get ReadOnlyProperty1() As String 'unreferenced
End Property

Public Property Let ReadOnlyProperty1(ByVal RHS As String) 'unreferenced
End Property

Public Property Get SomeOtherProperty() As String
End Property

Private Sub ReferencingSub(arg As String)
   Dim tester As String
   tester = ReadOnlyProperty
   ReadOnlyProperty = tester & arg & SomeOtherProperty
End Sub
";

            const string expectedCode =
@"
'Only calling 'ReferencingSub' here to give it a reference for the test
Public Property Get ReadOnlyProperty() As String
    ReferencingSub ""arg""
End Property

Public Property Let ReadOnlyProperty(ByVal RHS As String)
End Property

Public Property Get SomeOtherProperty() As String
End Property
";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ProcedureNotUsedInspection(state));
            StringAssert.StartsWith(expectedCode.Trim(), actualCode.Trim());
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ApplyToMultiple()
        {
            var inputCode =
@"
Option Explicit

Public notUsed1 As Long
Public notUsed2 As Long
Public notUsed3 As Long
";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new VariableNotUsedInspection(state));
            StringAssert.Contains("Option Explicit", actualCode);
            StringAssert.DoesNotContain("notUsed", actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(RemoveUnusedDeclarationQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void ApplyToList()
        {
            var inputCode =
@"
Option Explicit

Public notUsed1 As Long, notUsed2 As Long, notUsed3 As Long
";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new VariableNotUsedInspection(state));
            StringAssert.Contains("Option Explicit", actualCode);
            StringAssert.DoesNotContain("notUsed", actualCode);
            StringAssert.DoesNotContain("Public", actualCode);
        }

        private string ApplyQuickFixToFirstInspectionResult(string inputCode, Func<RubberduckParserState, IInspection> inspectionFactory = null, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var resultToFix = inspectionResults.First();

                var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                    .Resolve<DeleteDeclarationsRefactoringAction>();

                var quickFix = new RemoveUnusedDeclarationQuickFix(deleteDeclarationRefactoringAction);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                quickFix.Fix(resultToFix, rewriteSession);
                rewriteSession.TryRewrite();

                return component.CodeModule.Content();
            }
        }

        private string ApplyQuickFixToAllInspectionResults(string inputCode, Func<RubberduckParserState, IInspection> inspectionFactory = null, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                    .Resolve<DeleteDeclarationsRefactoringAction>();
                
                var quickFix = new RemoveUnusedDeclarationQuickFix(deleteDeclarationRefactoringAction);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                quickFix.Fix(inspectionResults.ToList(), rewriteSession);
                rewriteSession.TryRewrite();

                return component.CodeModule.Content();
            }
        }
    }
}
