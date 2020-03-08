using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Logistics;
using Rubberduck.Parsing.Rewriter;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class QuickFixProviderTests
    {
        [Test]
        [Category("QuickFixes")]
        public void ProviderDoesNotKnowAboutInspection()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new ConstantNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                var failureNotifier = new Mock<IQuickFixFailureNotifier>().Object;
                var quickFixProvider = new QuickFixProvider(rewritingManager, failureNotifier, new IQuickFix[] { });
                Assert.AreEqual(0, quickFixProvider.QuickFixes(inspectionResults.First()).Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProviderKnowsAboutInspection()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim str As String
    str = """"
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new EmptyStringLiteralInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                var failureNotifier = new Mock<IQuickFixFailureNotifier>().Object;
                var quickFixProvider = new QuickFixProvider(rewritingManager, failureNotifier, new IQuickFix[] { new ReplaceEmptyStringLiteralStatementQuickFix() });
                Assert.AreEqual(1, quickFixProvider.QuickFixes(inspectionResults.First()).Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ResultDisablesFix()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new ConstantNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                var failureNotifier = new Mock<IQuickFixFailureNotifier>().Object;
                var quickFixProvider = new QuickFixProvider(rewritingManager, failureNotifier, new IQuickFix[] { new RemoveUnusedDeclarationQuickFix() });

                var result = inspectionResults.First();
                result.DisabledQuickFixes.Add(nameof(RemoveUnusedDeclarationQuickFix));

                Assert.AreEqual(0, quickFixProvider.QuickFixes(result).Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProviderCallsNotifierOnFailureToRewrite()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim str As String
    str = """"
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new EmptyStringLiteralInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                var inspectionResult = inspectionResults.First();

                var failureNotifierMock = new Mock<IQuickFixFailureNotifier>();
                var quickFixProvider = new QuickFixProvider(rewritingManager, failureNotifierMock.Object, new IQuickFix[] { new ReplaceEmptyStringLiteralStatementQuickFix() });
                var quickFix = quickFixProvider.QuickFixes(inspectionResult).First();

                //Make rewrite fail.
                component.CodeModule.InsertLines(1, "'afejfaofef");

                quickFixProvider.Fix(quickFix, inspectionResult);

                failureNotifierMock.Verify(m => m.NotifyQuickFixExecutionFailure(RewriteSessionState.StaleParseTree));
            }
        }
    }
}