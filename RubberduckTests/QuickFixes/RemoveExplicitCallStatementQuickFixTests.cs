using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveExplicitCallStatementQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCallStatement_QuickFixWorks_RemoveCallStatement()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Goo 1, ""test""
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteCallStatementInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveExplicitCallStatementQuickFix();
        }
    }
}
