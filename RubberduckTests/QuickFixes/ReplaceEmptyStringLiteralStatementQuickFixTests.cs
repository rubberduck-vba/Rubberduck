using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
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
