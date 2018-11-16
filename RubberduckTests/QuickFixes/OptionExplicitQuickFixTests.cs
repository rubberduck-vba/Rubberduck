using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class OptionExplicitQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void NotAlreadySpecified_QuickFixWorks()
        {
            const string inputCode = "";
            const string expectedCode =
@"Option Explicit

";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new OptionExplicitInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new OptionExplicitQuickFix();
        }
    }
}
