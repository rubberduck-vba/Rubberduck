using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveEmptyElseBlockQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void EmptyElseBlock_QuickFixRemovesElse()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
    End If
End Sub";

            const string expectedCode =
                @"Sub Foo()
    If True Then
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveEmptyElseBlockQuickFix(state).Fix(actualResults.First());

                string actualRewrite = state.GetRewriter(component).GetText();

                Assert.AreEqual(expectedCode, actualRewrite);
            }
        }
    }
}
