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
    public class RemoveExplicitLetStatementQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteLetStatement_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteLetStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitLetStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
