using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
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
            const string inputCode = @"
Public Sub Test() ' inspection won't yield any results if module is empty (#2621)
End Sub
";
            const string expectedCode = @"Option Explicit

Public Sub Test() ' inspection won't yield any results if module is empty (#2621)
End Sub
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
