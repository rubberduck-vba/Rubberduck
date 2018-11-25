using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class MakeSingleLineParamegerQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void MultilineParameter_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer)
End Sub";

            const string expectedCode =
                @"Public Sub Foo( _
    ByVal Var1 As Integer)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MultilineParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new MakeSingleLineParameterQuickFix();
        }
    }
}
