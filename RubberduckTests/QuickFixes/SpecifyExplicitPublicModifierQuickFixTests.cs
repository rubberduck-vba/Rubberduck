using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;


namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SpecifyExplicitPublicModifierQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ImplicitPublicMember_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 as Integer)
'Just an inoffensive little comment

End Sub";

            const string expectedCode =
                @"Public Sub Foo(ByVal arg1 as Integer)
'Just an inoffensive little comment

End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitPublicMemberInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new SpecifyExplicitPublicModifierQuickFix();
        }
    }
}
