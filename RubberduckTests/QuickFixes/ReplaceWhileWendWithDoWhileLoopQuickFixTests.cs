using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceWhileWendWithDoWhileLoopQuickFixTests : QuickFixTestBase
    {
        protected override IQuickFix QuickFix(RubberduckParserState state)
            => new ReplaceWhileWendWithDoWhileLoopQuickFix();

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteWhileWendStatement_QuickFixWorks()
        {
            const string input = @"
Sub Foo()
    While True
    Wend
End Sub
";
            const string expected = @"
Sub Foo()
    Do While True
    Loop
End Sub
";
            var actual = ApplyQuickFixToFirstInspectionResult(input, state => new ObsoleteWhileWendStatementInspection(state));
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteWhileWendStatement_InstructionsSeparator_QuickFixWorks()
        {
            const string input = @"
Sub Foo()
    While True : DoSomething : Wend
End Sub
";
            const string expected = @"
Sub Foo()
    Do While True : DoSomething : Loop
End Sub
";
            var actual = ApplyQuickFixToFirstInspectionResult(input, state => new ObsoleteWhileWendStatementInspection(state));
            Assert.AreEqual(expected, actual);
        }
    }
}