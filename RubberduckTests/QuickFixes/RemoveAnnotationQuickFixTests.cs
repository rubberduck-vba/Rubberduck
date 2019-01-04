using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveAnnotationQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ModuleAttributeAnnotationWithoutAttribute_QuickFixWorks()
        {
            const string inputCode =
                @"'@PredeclaredId
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MemberAttributeAnnotationWithoutAttribute_QuickFixWorks()
        {
            const string inputCode =
                @"'@DefaultMember
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveAnnotationQuickFix(new AnnotationUpdater());
        }
    }
}