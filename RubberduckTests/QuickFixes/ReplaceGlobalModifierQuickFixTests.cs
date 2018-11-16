using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
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
