using System.Linq;
using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveLocalErrorQuickFixTests
    {
        [TestCase("On Local Error GoTo 0")]
        [TestCase("On Local Error GoTo 1")]
        [TestCase(@"On Local Error GoTo Label
Label:")]
        [TestCase("On Local Error Resume Next")]
        [Category("QuickFixes")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement(string stmt)
        {
            var inputCode = $@"Sub foo()
    {stmt}
End Sub";
            var expectedCode = $@"Sub foo()
    {stmt.Replace("Local ", "")}
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new OnLocalErrorInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveLocalErrorQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
