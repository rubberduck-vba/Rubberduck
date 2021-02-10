using Moq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactoring.ParseTreeValue;
using Rubberduck.Refactorings.ImplicitTypeToExplicit;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using System.Threading;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class DeclareAsExplicitTypeQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_Variable()
        {
            const string inputCode =
@"Sub Foo(arg As String)
    Dim var1
    var1 = arg
End Sub";

            const string expectedCode =
@"Sub Foo(arg As String)
    Dim var1 As String
    var1 = arg
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_Parameter()
        {
            const string inputCode =
@"
Private mArg As Long

Sub Foo(arg1)
    arg1 = mArg
End Sub";

            const string expectedCode =
@"
Private mArg As Long

Sub Foo(arg1 As Long)
    arg1 = mArg
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }
        
        private string ApplyQuickFixToFirstInspectionResult(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var resultToFix = inspectionResults.First();

                var refactoringAction = new ImplicitTypeToExplicitRefactoringAction(state, new ParseTreeValueFactory(), rewritingManager);
                var quickFix = new DeclareAsExplicitTypeQuickFix(refactoringAction);

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();
                quickFix.Fix(resultToFix, rewriteSession);
                rewriteSession.TryRewrite();

                return component.CodeModule.Content();
            }
        }
    }
}
