using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;


namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class ReplaceGlobalModifierQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ObsoleteGlobal_QuickFixWorks()
        {
            const string inputCode =
                @"Global var1 As Integer";

            const string expectedCode =
                @"Public var1 As Integer";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ReplaceGlobalModifierQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
