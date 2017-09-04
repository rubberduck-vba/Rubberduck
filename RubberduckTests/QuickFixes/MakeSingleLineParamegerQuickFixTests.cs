using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class MakeSingleLineParamegerQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MultilineParameterInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new MakeSingleLineParameterQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

    }
}
