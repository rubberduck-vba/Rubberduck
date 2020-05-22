using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveCommentQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment()
        {
            const string inputCode =
                @"Rem test1";

            const string expectedCode =
                @"";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCode()
        {
            const string inputCode =
                @"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
                @"Dim Foo As Integer: ";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveCommentQuickFix();
        }
    }
}
