using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class ChangeDimToPrivateQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ModuleScopeDimKeyword_QuickFixWorks()
        {
            const string inputCode =
                @"Dim foo As String";

            const string expectedCode =
                @"Private foo As String";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ModuleScopeDimKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeDimToPrivateQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ModuleScopeDimKeyword_QuickFixWorks_SplitDeclaration()
        {
            const string inputCode =
                @"Dim _
      foo As String";

            const string expectedCode =
                @"Private _
      foo As String";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ModuleScopeDimKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeDimToPrivateQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ModuleScopeDimKeyword_QuickFixWorks_MultipleDeclarations()
        {
            const string inputCode =
                @"Dim foo As String, _
      bar As Integer";

            const string expectedCode =
                @"Private foo As String, _
      bar As Integer";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ModuleScopeDimKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeDimToPrivateQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
