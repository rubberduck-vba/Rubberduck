using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class MakeSingleLineParamegerQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void MultilineParameter_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer)
End Sub";

            const string expectedCode =
                @"Public Sub Foo( _
    ByVal Var1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new MakeSingleLineParameterQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
