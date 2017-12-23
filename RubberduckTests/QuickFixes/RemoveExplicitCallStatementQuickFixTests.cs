using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveExplicitCallStatementQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCallStatement_QuickFixWorks_RemoveCallStatement()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Goo 1, ""test""
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                var fix = new RemoveExplicitCallStatmentQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
