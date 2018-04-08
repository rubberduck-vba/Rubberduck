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
    public class ReplaceEmptyStringLiteralStatementQuickFixTests
    {
        [Test]
        [Category("Inspections")]
        public void EmptyStringLiteral_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
    arg1 = """"
End Sub";

            const string expectedCode =
                @"Public Sub Foo(ByRef arg1 As String)
    arg1 = vbNullString
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyStringLiteralInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceEmptyStringLiteralStatementQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
