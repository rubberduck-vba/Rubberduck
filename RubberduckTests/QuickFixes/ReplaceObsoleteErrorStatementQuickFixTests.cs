using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceObsoleteErrorStatementQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteErrorSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteErrorSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteErrorSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteErrorSyntaxInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ReplaceObsoleteErrorStatementQuickFix();
        }
    }
}
