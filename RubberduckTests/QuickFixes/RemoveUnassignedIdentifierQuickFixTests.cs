using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.DeleteDeclarations;
using RubberduckTests.Mocks;
using RubberduckTests.Refactoring.DeleteDeclarations;
using Rubberduck.Parsing.Rewriter;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnassignedIdentifierQuickFixTests
    {
        [Test]
        [Category(nameof(RemoveUnassignedIdentifierQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Category("QuickFixes")]
        public void UnassignedVariable_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 as Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotAssignedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category(nameof(RemoveUnassignedIdentifierQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Category("QuickFixes")]
        public void UnassignedVariable_VariableOnMultipleLines_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim _
var1 _
as _
Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotAssignedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category(nameof(RemoveUnassignedIdentifierQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Category("QuickFixes")]
        public void UnassignedVariable_MultipleVariablesOnSingleLine_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As Integer, var2 As Boolean
End Sub";

            const string expectedCode =
                @"Sub Foo()
Dim var1 As Integer
End Sub";
            Func<IInspectionResult, bool> conditionToFix = s => s.Target.IdentifierName == "var2";
            var actualCode = ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(inputCode, state => new VariableNotAssignedInspection(state), conditionToFix);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category(nameof(RemoveUnassignedIdentifierQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Category("QuickFixes")]
        public void UnassignedVariable_MultipleVariablesOnMultipleLines_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As Integer, _
var2 As Boolean
End Sub";

            const string expectedCode =
                @"Sub Foo()
Dim var1 As Integer
End Sub";

            Func<IInspectionResult, bool> conditionToFix = result => result.Target.IdentifierName == "var2";
            var actualCode = ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(inputCode, state => new VariableNotAssignedInspection(state), conditionToFix);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category(nameof(RemoveUnassignedIdentifierQuickFix))]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        [Category("QuickFixes")]
        public void UnassignedVariable_MultipleVariablesOnMultipleLines_FixManyWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As Integer, _
var2 As Boolean, var3 as Long, var4 As String
End Sub";

            const string expectedCode =
                @"Sub Foo()
Dim var1 As Integer, _
var3 as Long
End Sub";
            var targets = new string[] { "var2", "var4" };
            Func<IInspectionResult, bool> conditionToFix = result => targets.Contains(result.Target.IdentifierName);
            var actualCode = ApplyQuickFixToAllInspectionResultsSatisfyingPredicate(inputCode, state => new VariableNotAssignedInspection(state), conditionToFix);
            Assert.AreEqual(expectedCode, actualCode);
        }

        private string ApplyQuickFixToFirstInspectionResult(string inputCode, Func<RubberduckParserState, IInspection> inspectionFactory = null, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            return ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(
                inputCode, inspectionFactory, 
                (result) => result != null);            
        }

        private string ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(string inputCode, Func<RubberduckParserState, IInspection> inspectionFactory = null, Func<IInspectionResult, bool> predicate = null, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var (inspectionResults, quickFix) = GetInspectionResultsAndQuickFix(inputCode, state, rewritingManager, inspectionFactory);

                var rewriteSession = codeKind == CodeKind.AttributesCode
                    ? rewritingManager.CheckOutAttributesSession()
                    : rewritingManager.CheckOutCodePaneSession();

                quickFix.Fix(inspectionResults.First(predicate), rewriteSession);

                rewriteSession.TryRewrite();

                return component.CodeModule.Content();
            }
        }

        private string ApplyQuickFixToAllInspectionResultsSatisfyingPredicate(string inputCode, Func<RubberduckParserState, IInspection> inspectionFactory = null, Func<IInspectionResult, bool> predicate = null, CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var (inspectionResults, quickFix) = GetInspectionResultsAndQuickFix(inputCode, state, rewritingManager, inspectionFactory);
                
                var rewriteSession = codeKind == CodeKind.AttributesCode
                    ? rewritingManager.CheckOutAttributesSession()
                    : rewritingManager.CheckOutCodePaneSession();

                quickFix.Fix(inspectionResults.Where(r => predicate(r)).ToList(), rewriteSession);

                rewriteSession.TryRewrite();

                return component.CodeModule.Content();
            }
        }

        private (IEnumerable<IInspectionResult> InspResults, IQuickFix QFix) GetInspectionResultsAndQuickFix(string inputCode, RubberduckParserState state, IRewritingManager rewritingManager, Func<RubberduckParserState, IInspection> inspectionFactory = null)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var inspection = inspectionFactory(state);
            var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

            var deleteDeclarationRefactoringAction = new DeleteDeclarationsTestsResolver(state, rewritingManager)
                .Resolve<DeleteDeclarationsRefactoringAction>();

            var quickFix = new RemoveUnassignedIdentifierQuickFix(deleteDeclarationRefactoringAction);

            return (inspectionResults, quickFix);
        }
    }
}
