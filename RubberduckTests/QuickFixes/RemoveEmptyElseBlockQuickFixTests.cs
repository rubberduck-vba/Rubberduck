using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
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
