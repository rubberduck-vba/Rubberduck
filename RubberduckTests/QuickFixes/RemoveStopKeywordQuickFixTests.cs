using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class RemoveStopKeywordQuickFixTests
    {

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void StopKeyword_QuickFixWorks_RemoveKeyword()
        {
            var inputCode =
                @"Sub Foo()
    Stop
End Sub";

            var expectedCode =
                @"Sub Foo()
    
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new StopKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveStopKeywordQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void StopKeyword_QuickFixWorks_RemoveKeyword_InstructionSeparator()
        {
            var inputCode = "Sub Foo(): Stop: End Sub";

            var expectedCode = "Sub Foo(): : End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new StopKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveStopKeywordQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
