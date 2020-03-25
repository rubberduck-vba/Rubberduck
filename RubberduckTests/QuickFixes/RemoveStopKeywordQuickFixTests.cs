using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveStopKeywordQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void StopKeyword_QuickFixWorks_RemoveKeyword()
        {
            var inputCode =
                @"Sub Foo()
    Stop
End Sub";

            var expectedCode =
                @"Sub Foo()
    
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new StopKeywordInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void StopKeyword_QuickFixWorks_RemoveKeyword_InstructionSeparator()
        {
            var inputCode = "Sub Foo(): Stop: End Sub";

            var expectedCode = "Sub Foo(): : End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new StopKeywordInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveStopKeywordQuickFix();
        }
    }
}
