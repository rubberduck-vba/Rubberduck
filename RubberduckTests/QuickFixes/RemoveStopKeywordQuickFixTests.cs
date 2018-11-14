using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
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
