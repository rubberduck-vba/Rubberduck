﻿using System.Linq;
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
    public class ReplaceOboleteCommentMarkerQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"' test1";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateCommentHasContinuation()
        {
            const string inputCode =
@"Rem this is _
a comment";

            const string expectedCode =
@"' this is _
a comment";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }


        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCode()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
@"Dim Foo As Integer: ' This is a comment";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
@"Dim Foo As Integer: ' This is _
a comment";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }
    }
}
