using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceObsoleteCommentMarkerQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment()
        {
            const string inputCode =
                @"Rem test1";

            const string expectedCode =
                @"' test1";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode); 
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateCommentHasContinuation()
        {
            const string inputCode =
                @"Rem this is _
a comment";

            const string expectedCode =
                @"' this is _
a comment";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCode()
        {
            const string inputCode =
                @"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
                @"Dim Foo As Integer: ' This is a comment";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
                @"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
                @"Dim Foo As Integer: ' This is _
a comment";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ReplaceObsoleteCommentMarkerQuickFix();
        }
    }
}
