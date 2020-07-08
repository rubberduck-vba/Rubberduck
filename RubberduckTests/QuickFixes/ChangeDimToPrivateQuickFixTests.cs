using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ChangeDimToPrivateQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ModuleScopeDimKeyword_QuickFixWorks()
        {
            const string inputCode =
                @"Dim foo As String";

            const string expectedCode =
                @"Private foo As String";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ModuleScopeDimKeywordInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ModuleScopeDimKeyword_QuickFixWorks_SplitDeclaration()
        {
            const string inputCode =
                @"Dim _
      foo As String";

            const string expectedCode =
                @"Private _
      foo As String";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ModuleScopeDimKeywordInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ModuleScopeDimKeyword_QuickFixWorks_MultipleDeclarations()
        {
            const string inputCode =
                @"Dim foo As String, _
      bar As Integer";

            const string expectedCode =
                @"Private foo As String, _
      bar As Integer";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ModuleScopeDimKeywordInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ChangeDimToPrivateQuickFix();
        }
    }
}
