using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveEmptyElseBlockQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void EmptyElseBlock_QuickFixRemovesElse()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
    End If
End Sub";

            const string expectedCode =
                @"Sub Foo()
    If True Then
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyElseBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveEmptyElseBlockQuickFix();
        }
    }
}
