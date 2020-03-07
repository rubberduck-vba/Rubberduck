using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;


namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceGlobalModifierQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ObsoleteGlobal_QuickFixWorks()
        {
            const string inputCode =
                @"Global var1 As Integer";

            const string expectedCode =
                @"Public var1 As Integer";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteGlobalInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ReplaceGlobalModifierQuickFix();
        }
    }
}
