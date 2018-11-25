using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
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

                var quickFixProvider = new QuickFixProvider(rewritingManager, new IQuickFix[] { });
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

                var quickFixProvider = new QuickFixProvider(rewritingManager, new IQuickFix[] { new ReplaceEmptyStringLiteralStatementQuickFix() });
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

                var quickFixProvider = new QuickFixProvider(rewritingManager, new IQuickFix[] { new RemoveUnusedDeclarationQuickFix() });

                var result = inspectionResults.First();
                result.Properties.DisableFixes = nameof(RemoveUnusedDeclarationQuickFix);

                Assert.AreEqual(0, quickFixProvider.QuickFixes(result).Count());
            }
        }
    }
}