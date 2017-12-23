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
    public class ReplaceObsoleteErrorStatementQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Error 91
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Err.Raise 91
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteErrorSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceObsoleteErrorStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_ProcNamedError()
        {
            const string inputCode =
                @"Sub Error(val as Integer)
End Sub

Sub Foo()
    Error 91
End Sub";

            const string expectedCode =
                @"Sub Error(val as Integer)
End Sub

Sub Foo()
    Err.Raise 91
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteErrorSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceObsoleteErrorStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateCommentHasContinuation()
        {
            const string inputCode =
                @"Sub Foo()
    Error _
    91
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Err.Raise _
    91
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteErrorSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceObsoleteErrorStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }


        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCode()
        {
            const string inputCode =
                @"Sub Foo()
    Dim foo: Error 91
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim foo: Err.Raise 91
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteErrorSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceObsoleteErrorStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
