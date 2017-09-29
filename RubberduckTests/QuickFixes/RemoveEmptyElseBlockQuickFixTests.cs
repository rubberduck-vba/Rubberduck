using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class RemoveEmptyElseBlockQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new EmptyElseBlockInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveEmptyElseBlockQuickFix(state).Fix(actualResults.First());

            string actualRewrite = state.GetRewriter(component).GetText();

            Assert.AreEqual(expectedCode, actualRewrite);
        }
    }
}
