using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SplitMultipleDeclarationsQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_QuickFixWorks_Variables()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Dim var1 As Integer
Dim var2 As String

End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MultipleDeclarationsInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_QuickFixWorks_Constants()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const var1 As Integer = 9, var2 As String = ""test""
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Const var1 As Integer = 9
Const var2 As String = ""test""

End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MultipleDeclarationsInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_QuickFixWorks_StaticVariables()
        {
            const string inputCode =
                @"Public Sub Foo()
    Static var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Static var1 As Integer
Static var2 As String

End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MultipleDeclarationsInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new SplitMultipleDeclarationsQuickFix();
        }
    }
}
