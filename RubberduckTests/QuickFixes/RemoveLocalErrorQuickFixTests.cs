using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveLocalErrorQuickFixTests : QuickFixTestBase
    {
        [TestCase("On Local Error GoTo 0")]
        [TestCase("On Local Error GoTo 1")]
        [TestCase(@"On Local Error GoTo Label
Label:")]
        [TestCase("On Local Error Resume Next")]
        [Category("QuickFixes")]
        public void OptionBaseZeroStatement_QuickFixWorks_RemoveStatement(string stmt)
        {
            var inputCode = $@"Sub foo()
    {stmt}
End Sub";
            var expectedCode = $@"Sub foo()
    {stmt.Replace("Local ", "")}
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new OnLocalErrorInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveLocalErrorQuickFix();
        }
    }
}
