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
    public class RemoveCommentQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment()
        {
            const string inputCode =
                @"Rem test1";

            const string expectedCode =
                @"";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveCommentHasContinuation()
        {
            const string inputCode =
                @"Rem test1 _
continued";

            const string expectedCode =
                @"";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCode()
        {
            const string inputCode =
                @"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
                @"Dim Foo As Integer: ";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
                @"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
                @"Dim Foo As Integer: ";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
