using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceEmptyStringLiteralStatementQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyStringLiteralInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ReplaceEmptyStringLiteralStatementQuickFix();
        }
    }
}
