using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class QuickFixProviderTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ProviderDoesNotKnowAboutInspection()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ConstantNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var quickFixProvider = new QuickFixProvider(state, new IQuickFix[] { });
                Assert.AreEqual(0, quickFixProvider.QuickFixes(inspectionResults.First()).Count());
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ProviderKnowsAboutInspection()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim str As String
    str = """"
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyStringLiteralInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                var quickFixProvider = new QuickFixProvider(state, new IQuickFix[] { new ReplaceEmptyStringLiteralStatementQuickFix(state) });
                Assert.AreEqual(1, quickFixProvider.QuickFixes(inspectionResults.First()).Count());
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ResultDisablesFix()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ConstantNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var quickFixProvider = new QuickFixProvider(state, new IQuickFix[] { new RemoveUnusedDeclarationQuickFix(state) });

                var result = inspectionResults.First();
                result.Properties.Add("DisableFixes", nameof(RemoveUnusedDeclarationQuickFix));

                Assert.AreEqual(0, quickFixProvider.QuickFixes(result).Count());
            }
        }
    }
}