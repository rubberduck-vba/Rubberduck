using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveExplicitLetStatementQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteLetStatement_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObsoleteLetStatementInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveExplicitLetStatementQuickFix();
        }
    }
}
